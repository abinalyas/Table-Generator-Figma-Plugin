import { fakerMethods } from './faker-methods';

interface ComponentInfo {
  id: string | null;
  name: string;
  width?: number;
  properties: { [key: string]: any };
  availableProperties: string[];
  propertyTypes: { [key: string]: any };
}

interface State {
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

interface Elements {
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
  componentModeBtn: HTMLButtonElement;
  propertyEditorOverlay: HTMLElement;
  componentDisplay: HTMLElement;
  scanOptionsContainer: HTMLElement | null;
  scanTableBtn?: HTMLButtonElement;
  scanTableSection?: HTMLElement;
}



// State Management
const state: State = {
  selectedCells: new Set<string>(),
  isDragging: false,
  startCell: null,
  endCell: null,
  componentProps: {},
  cellProperties: new Map<string, any>(),
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
  footerComponent: null,
  tableFrameId: undefined,
};

// Initialize the tooltip element
state.tooltip.className = 'cell-tooltip';
state.tooltip.style.position = 'fixed';
state.tooltip.style.display = 'none';
state.tooltip.style.background = 'rgba(0, 0, 0, 0.85)';
state.tooltip.style.color = 'white';
state.tooltip.style.padding = '8px 12px';
state.tooltip.style.borderRadius = '4px';
state.tooltip.style.fontSize = '12px';
state.tooltip.style.zIndex = '10000';
state.tooltip.style.maxWidth = '300px';
state.tooltip.style.wordWrap = 'break-word';
state.tooltip.style.pointerEvents = 'none';

// DOM Elements
let elements = {} as Elements;

// DOM Elements for faker dropdown
const fakerMethodInput = document.getElementById('fakerMethodInput') as HTMLInputElement;
const fakerDropdown = document.getElementById('fakerDropdown')!;

// Create tooltip element has been moved to state initialization

let domReady = false;
// Ensure property editor scaffold exists
function ensurePropertyEditorScaffold() {
  try {
    const editor = elements.propertyEditor;
    if (!editor) return;
    // Create dynamicPropertyFields if missing
    let dyn = document.getElementById('dynamicPropertyFields');
    if (!dyn) {
      dyn = document.createElement('div');
      dyn.id = 'dynamicPropertyFields';
      editor.insertBefore(dyn, editor.querySelector('.property-actions'));
      console.log('[DEBUG] Created #dynamicPropertyFields at runtime');
    }
    // Create colWidthContainer + input if missing
    let colWrap = document.getElementById('colWidthContainer');
    if (!colWrap) {
      colWrap = document.createElement('div');
      colWrap.id = 'colWidthContainer';
      colWrap.className = 'property-field';
      colWrap.style.display = 'none';
      const label = document.createElement('label');
      label.htmlFor = 'colWidthInput';
      label.textContent = 'Column width';
      const input = document.createElement('input');
      input.type = 'number';
      input.id = 'colWidthInput';
      input.className = 'styled-input';
      input.min = '1';
      input.placeholder = 'Enter width';
      colWrap.appendChild(label);
      colWrap.appendChild(input);
      editor.insertBefore(colWrap, editor.querySelector('.property-actions'));
      // refresh element refs
      elements.colWidthContainer = colWrap as HTMLElement;
      elements.colWidthInput = input as HTMLInputElement;
      console.log('[DEBUG] Created #colWidthContainer and #colWidthInput at runtime');
    } else {
      // Refresh references in case they were null
      const input = document.getElementById('colWidthInput') as HTMLInputElement | null;
      if (input) elements.colWidthInput = input;
      elements.colWidthContainer = colWrap as HTMLElement;
    }
  } catch (e) {
    console.error('[DEBUG] ensurePropertyEditorScaffold failed:', e);
  }
}

// Static property models for header and footer
const HEADER_CELL_MODEL = [
  { name: "Cell text#12234:32", label: "Header Text", type: "TEXT" },
  { name: "State", label: "State", type: "VARIANT", options: ["Enable", "Hover", "Focus"] },
  { name: "Sortable", label: "Sortable", type: "VARIANT", options: ["True", "False"], defaultValue: "False" },
  { name: "Sorted", label: "Sorted", type: "VARIANT", options: ["Ascending", "Descending", "None"], defaultValue: "Ascending", dependsOn: "Sortable", showWhen: "True" }
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
  elements.statusMessage = document.getElementById('statusMessage')!;
  elements.gridContainer = document.getElementById('gridContainer')!;
  elements.actionButtons = document.getElementById('actionButtons')!;
  elements.createTableBtn = document.getElementById('createTableBtn') as HTMLButtonElement;
  elements.clearSelectionBtn = document.getElementById('clearSelectionBtn') as HTMLButtonElement;
  elements.propertyEditor = document.getElementById('propertyEditor')!;
  elements.propertyEditorTitle = document.getElementById('propertyEditorTitle')!;
  elements.editingCellCoords = document.getElementById('editingCellCoords')!;
  elements.showText = document.getElementById('showText') as HTMLInputElement;
  elements.secondTextLine = document.getElementById('secondTextLine') as HTMLInputElement;
  elements.secondCellText = document.getElementById('secondCellText') as HTMLInputElement;
  elements.state = document.getElementById('state') as HTMLSelectElement;
  elements.cancelPropsBtn = document.getElementById('cancelPropsBtn') as HTMLButtonElement;
  elements.savePropsBtn = document.getElementById('savePropsBtn') as HTMLButtonElement;
  elements.secondLineContainer = document.getElementById('secondLineContainer')!;
  elements.secondCellTextContainer = document.getElementById('secondCellTextContainer')!;
  elements.loader = document.getElementById('loader')!;
  elements.loaderText = document.getElementById('loaderText')!;
  elements.aiConfirmationDialog = document.getElementById('aiConfirmationDialog')!;
  elements.cancelAiBtn = document.getElementById('cancelAiBtn') as HTMLButtonElement;
  elements.confirmAiBtn = document.getElementById('confirmAiBtn') as HTMLButtonElement;
  elements.generateSampleCheckbox = document.getElementById('generateSampleCheckbox') as HTMLInputElement;
  elements.aiPromptContainer = document.getElementById('aiPromptContainer')!;
  elements.colWidthInput = document.getElementById('colWidthInput') as HTMLInputElement;
  elements.colWidthContainer = document.getElementById('colWidthContainer')!;
  elements.slotCheckbox = document.getElementById('slotCheckbox') as HTMLInputElement;
  elements.customCellTextToggle = document.getElementById('customCellTextToggle') as HTMLInputElement;
  elements.customCellText = document.getElementById('customCellText') as HTMLInputElement;
  elements.customCellTextContainer = document.getElementById('customCellTextContainer')!;
  elements.landingPage = document.getElementById('landingPage')!;
  elements.componentModeBtn = document.getElementById('componentModeBtn') as HTMLButtonElement;
  elements.propertyEditorOverlay = document.getElementById('propertyEditorOverlay')!;
  elements.componentDisplay = document.getElementById('componentDisplay')!;
  elements.scanOptionsContainer = null;

  domReady = true;

  // Initialize reset button as disabled
  updateResetButtonState();

  // Ask the plugin to scan the current selection immediately
  parent.postMessage({ pluginMessage: { type: 'scan-table' } }, '*');

  // Now safe to call setup functions
  createGrid();

  // Wire main-page single prompt controls
  const useSinglePrompt = document.getElementById('useSinglePrompt') as HTMLInputElement | null;
  const singlePromptContainer = document.getElementById('singlePromptContainer') as HTMLElement | null;
  const generateTableFromPromptBtn = document.getElementById('generateTableFromPromptBtn') as HTMLButtonElement | null;
  useSinglePrompt?.addEventListener('change', () => {
    if (singlePromptContainer) {
      singlePromptContainer.style.display = useSinglePrompt.checked ? 'block' : 'none';
    }

    // If single prompt is selected, deselect file upload
    const useFileUpload = document.getElementById('useFileUpload') as HTMLInputElement | null;
    const fileUploadContainer = document.getElementById('fileUploadContainer') as HTMLElement | null;
    if (useSinglePrompt?.checked && useFileUpload) {
      useFileUpload.checked = false;
      if (fileUploadContainer) {
        fileUploadContainer.style.display = 'none';
      }
    }
  });

  // Enable/disable AI button based on prompt input
  const singlePromptText = document.getElementById('singlePromptText') as HTMLTextAreaElement | null;

  function updateAIButtonState() {
    if (generateTableFromPromptBtn && singlePromptText) {
      const hasPrompt = singlePromptText.value.trim().length > 0;
      generateTableFromPromptBtn.disabled = !hasPrompt;
      generateTableFromPromptBtn.style.opacity = hasPrompt ? '1' : '0.5';
      generateTableFromPromptBtn.style.cursor = hasPrompt ? 'pointer' : 'not-allowed';
    }
  }

  // Initially disable the button
  if (generateTableFromPromptBtn) {
    generateTableFromPromptBtn.disabled = true;
    generateTableFromPromptBtn.style.opacity = '0.5';
    generateTableFromPromptBtn.style.cursor = 'not-allowed';
  }

  // Listen for prompt input changes
  singlePromptText?.addEventListener('input', updateAIButtonState);
  singlePromptText?.addEventListener('paste', () => {
    // Use setTimeout to ensure paste content is processed
    setTimeout(updateAIButtonState, 10);
  });

  // Wire file upload controls
  const useFileUpload = document.getElementById('useFileUpload') as HTMLInputElement | null;
  const fileUploadContainer = document.getElementById('fileUploadContainer') as HTMLElement | null;
  useFileUpload?.addEventListener('change', () => {
    if (fileUploadContainer) {
      fileUploadContainer.style.display = useFileUpload.checked ? 'block' : 'none';
    }

    // If file upload is selected, deselect single prompt
    const useSinglePrompt = document.getElementById('useSinglePrompt') as HTMLInputElement | null;
    const singlePromptContainer = document.getElementById('singlePromptContainer') as HTMLElement | null;
    if (useFileUpload?.checked && useSinglePrompt) {
      useSinglePrompt.checked = false;
      if (singlePromptContainer) {
        singlePromptContainer.style.display = 'none';
      }
    }
  });

  // Handle file selection
  const dataFileInput = document.getElementById('dataFileInput') as HTMLInputElement | null;
  dataFileInput?.addEventListener('change', handleFileUpload);

  // Show data in grid preview
  function showDataInGrid(data: any[][]) {
    console.log('File data received:', data);
    // For now, just show in file preview
    showFilePreview(data);
  }

  // Handle file upload
  async function handleFileUpload(event: Event) {
    const input = event.target as HTMLInputElement;
    if (!input.files || input.files.length === 0) {
      return;
    }

    const file = input.files[0];

    try {
      showLoader('Parsing file...');

      let data: any[][] = [];

      if (file.name.endsWith('.csv')) {
        data = await parseCSV(file);
      } else if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
        data = await parseExcel(file);
      } else if (file.name.endsWith('.json')) {
        data = await parseJSON(file);
      } else {
        throw new Error('Unsupported file format');
      }

      hideLoader();

      if (data.length === 0) {
        showMessage('File is empty or could not be parsed', 'error');
        return;
      }

      // Show data in grid preview
      showDataInGrid(data);

    } catch (error: any) {
      hideLoader();
      console.error('Error parsing file:', error);
      showMessage('Error parsing file: ' + error.message, 'error');
    }
  }

  generateTableFromPromptBtn?.addEventListener('click', () => {
    const rows = parseInt((document.getElementById('scanRowsInput') as HTMLInputElement)?.value || String(state.gridRows), 10);
    const cols = parseInt((document.getElementById('scanColsInput') as HTMLInputElement)?.value || String(state.gridCols), 10);
    const includeHeader = (document.getElementById('scanHeaderToggle') as HTMLInputElement)?.checked;
    const includeFooter = (document.getElementById('scanFooterToggle') as HTMLInputElement)?.checked;
    const includeSelectable = (document.getElementById('scanSelectableToggle') as HTMLInputElement)?.checked;
    const includeExpandable = (document.getElementById('scanExpandableToggle') as HTMLInputElement)?.checked;

    const prompt = (document.getElementById('singlePromptText') as HTMLTextAreaElement | null)?.value?.trim();
    if (!prompt) { showMessage('Enter a prompt', 'error'); return; }

    // Show loader when starting AI generation
    showLoader('Generating content with AI...');

    // Only using watsonx.ai for single prompt generation
    const apiKey = ((document.getElementById('spApiKey') as HTMLInputElement | null)?.value || '').trim();
    if (!apiKey) {
      hideLoader(); // Hide loader if there's an error
      showMessage('Provide IBM Cloud API key', 'error');
      return;
    }

    // Debug: Log what we're sending to the backend
    console.log('=== DEBUG: UI sending generate-table-with-ai message ===');
    console.log('Prompt:', prompt);
    console.log('Rows:', rows);
    console.log('Cols:', cols);
    console.log('API Key length:', apiKey.length);
    console.log('=== END DEBUG ===');

    parent.postMessage({ pluginMessage: { type: 'generate-table-with-ai', prompt, apiKey, rows, cols } }, '*');
  });
  setupEventListeners();
  setupResizeCorner();
  setupFileUploadControls();
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

  // Initialize tooltip
  state.tooltip.className = 'cell-tooltip';

});

function renderHeaderFooterGrids() {
  const headerGrid = document.getElementById('headerGrid');
  if (headerGrid) {
    headerGrid.innerHTML = '';
    headerGrid.style.setProperty('--header-cols', String(state.gridCols));
    
    // Apply sorting to maintain consistency with body cells
    const sortedCellProperties = sortColumnData(state.cellProperties, state.gridCols, state.gridRows);
    
    for (let c = 1; c <= state.gridCols; c++) {
      const cell = document.createElement('div');
      cell.className = 'header-cell';
      cell.dataset.col = String(c);

      const headerKey = `header-${c}`;
      const headerState = sortedCellProperties.get(headerKey);
      let displayText = '';
      
      if (headerState?.properties) {
        const textProp = state.headerCellComponent?.availableProperties?.find(p =>
          state.headerCellComponent!.propertyTypes[p] === 'TEXT');
        displayText = (textProp && headerState.properties[textProp]) || 
                     headerState.properties['Cell text#12234:32'] || '';
      }

      cell.textContent = displayText || `H${c}`;
      if (displayText) {
        cell.classList.add('edited');
        cell.style.fontWeight = 'bold';
      }

      cell.addEventListener('click', () => openPropertyEditor(`header-${c}`));
      cell.addEventListener('mouseenter', (e) => showCellTooltip(e.target as HTMLDivElement));
      cell.addEventListener('mouseleave', hideCellTooltip);
      headerGrid.appendChild(cell);
    }
  }
  
  const footerGrid = document.getElementById('footerGrid');
  if (footerGrid) {
    footerGrid.innerHTML = '';
    const cell = document.createElement('div');
    cell.className = 'footer-cell';
    cell.style.gridColumn = `span ${state.gridCols}`;

    const footerState = state.cellProperties.get('footer');
    let displayText = '';
    
    if (footerState?.properties) {
      const textProp = state.footerComponent?.availableProperties?.find(p =>
        state.footerComponent!.propertyTypes[p] === 'TEXT');
      displayText = (textProp && footerState.properties[textProp]) || 
                   footerState.properties['Total items#12006:49'] || '';
    }

    cell.textContent = displayText || 'Footer';
    if (displayText) {
      cell.classList.add('edited');
      cell.style.fontWeight = 'bold';
    }

    cell.addEventListener('click', () => openPropertyEditor('footer'));
    cell.addEventListener('mouseenter', (e) => showCellTooltip(e.target as HTMLDivElement));
    cell.addEventListener('mouseleave', hideCellTooltip);
    footerGrid.appendChild(cell);
  }
}

// Initialize
function createGrid() {
  const grid = elements.grid;
  grid.innerHTML = '';
  grid.style.gridTemplateColumns = `repeat(${state.gridCols}, 50px)`;

  const gridContainer = document.getElementById('gridContainer');
  if (gridContainer) {
    // Calculate total width: cellWidth(50px) * cols + gap(8px) * (cols-1) + padding(48px)
    const totalWidth = (state.gridCols * 50) + ((state.gridCols - 1) * 8) + 48;
    gridContainer.style.minWidth = `${Math.max(totalWidth, 400)}px`;
  }

  const sortedCellProperties = sortColumnData(state.cellProperties, state.gridCols, state.gridRows);
  // Update state with sorted data to ensure consistency
  state.cellProperties = sortedCellProperties;

  for (let r = 1; r <= state.gridRows; r++) {
    for (let c = 1; c <= state.gridCols; c++) {
      const cell = document.createElement('div');
      cell.className = 'cell';
      cell.dataset.row = String(r);
      cell.dataset.col = String(c);

      const key = `${r},${c}`;
      const cellState = sortedCellProperties.get(key);
      let displayText = '';
      let hasSlotEnabled = false;
      
      if (cellState?.properties && state.selectedComponent) {
        const slotProp = state.selectedComponent.availableProperties?.find(p =>
          state.selectedComponent!.propertyTypes[p] === 'BOOLEAN' && p.toLowerCase().includes('slot'));
        if (slotProp && cellState.properties[slotProp]) {
          hasSlotEnabled = true;
          cell.classList.add('slot-enabled');
        }

        const textProp = state.selectedComponent.availableProperties?.find(p =>
          state.selectedComponent!.propertyTypes[p] === 'TEXT');
        displayText = (textProp && cellState.properties[textProp]) || 
                     cellState.properties['Cell text#12234:32'] || '';
      }

      if (displayText) {
        cell.textContent = displayText;
        cell.classList.add('edited');
        cell.style.fontWeight = 'bold';
      } else {
        cell.textContent = hasSlotEnabled ? '' : `${r},${c}`;
      }

      cell.addEventListener('click', () => openPropertyEditor(key));
      cell.addEventListener('mouseenter', (e) => showCellTooltip(e.target as HTMLDivElement));
      cell.addEventListener('mouseleave', hideCellTooltip);
      grid.appendChild(cell);
    }
  }
  renderHeaderFooterGrids();
  updateCreateButtonState();
}



// Handle apply option clicks
function handleApplyOptionClick(this: HTMLElement) {
  console.log(`[DEBUG] Apply option clicked: ${this.dataset.apply}`);
  console.log(`[DEBUG] Current element:`, this);
  console.log(`[DEBUG] Element dataset:`, this.dataset);
  console.log(`[DEBUG] Element text content:`, this.textContent);
  console.log(`[DEBUG] Element classes before:`, this.className);

  // Remove active class from all options
  document.querySelectorAll('.apply-option').forEach(opt => {
    opt.classList.remove('active');
    console.log(`[DEBUG] Removed active from:`, opt.textContent);
  });

  // Add active class to clicked option
  this.classList.add('active');
  console.log(`[DEBUG] Added active to:`, this.textContent);
  console.log(`[DEBUG] Element classes after:`, this.className);
  // Update state
  const newApplyMode = this.dataset.apply as 'cell' | 'row' | 'column';
  state.applyMode = newApplyMode;

  // Show column width option only for cell and column apply modes
  if (state.applyMode === 'cell' || state.applyMode === 'column') {
    elements.colWidthContainer.style.display = 'block';
    console.log(`[DEBUG] Showing column width option for ${state.applyMode} mode`);
  } else {
    elements.colWidthContainer.style.display = 'none';
    console.log(`[DEBUG] Hiding column width option for ${state.applyMode} mode`);
  }
}

function setupApiKeySync() {
  const spApiKeyInput = document.getElementById('spApiKey') as HTMLInputElement | null;
  const watsonxApiKeyInput = document.getElementById('watsonxApiKey') as HTMLInputElement | null;

  if (!spApiKeyInput || !watsonxApiKeyInput) {
    console.warn('API key inputs not found');
    return;
  }

  // Function to sync API key between both inputs
  const syncApiKey = (sourceInput: HTMLInputElement, targetInput: HTMLInputElement, apiKey: string) => {
    if (apiKey.trim() && targetInput.value.trim() === '') {
      targetInput.value = apiKey;
      // Save the API key
      parent.postMessage({
        pluginMessage: {
          type: 'save-watsonx-api-key',
          apiKey: apiKey
        }
      }, '*');
    }
  };

  // Sync from single prompt to watsonx cell generation
  spApiKeyInput.addEventListener('input', () => {
    const apiKey = spApiKeyInput.value;
    syncApiKey(spApiKeyInput, watsonxApiKeyInput, apiKey);
  });

  // Sync from watsonx cell generation to single prompt
  watsonxApiKeyInput.addEventListener('input', () => {
    const apiKey = watsonxApiKeyInput.value;
    syncApiKey(watsonxApiKeyInput, spApiKeyInput, apiKey);
  });

  // Also sync on blur (when user finishes typing)
  spApiKeyInput.addEventListener('blur', () => {
    const apiKey = spApiKeyInput.value;
    if (apiKey.trim()) {
      watsonxApiKeyInput.value = apiKey;
      parent.postMessage({
        pluginMessage: {
          type: 'save-watsonx-api-key',
          apiKey: apiKey
        }
      }, '*');
    }
  });

  watsonxApiKeyInput.addEventListener('blur', () => {
    const apiKey = watsonxApiKeyInput.value;
    if (apiKey.trim()) {
      spApiKeyInput.value = apiKey;
      parent.postMessage({
        pluginMessage: {
          type: 'save-watsonx-api-key',
          apiKey: apiKey
        }
      }, '*');
    }
  });
}

function setupFileUploadControls() {
  const removeFileBtn = document.getElementById('removeFileBtn');

  if (removeFileBtn) {
    removeFileBtn.addEventListener('click', () => {
      // Clear the file input
      const fileInput = document.getElementById('dataFileInput') as HTMLInputElement | null;
      if (fileInput) {
        fileInput.value = '';
      }

      // Hide the remove button
      removeFileBtn.style.display = 'none';

      // Reset grid to default state
      resetGridToDefault();
    });
  }
}

function setupResizeCorner() {
  const resizeCorner = document.getElementById('resizeCorner') as HTMLElement;

  if (!resizeCorner) {
    console.warn('Resize corner element not found');
    return;
  }

  let isResizing = false;

  const resizeWindow = (e: PointerEvent) => {
    if (!isResizing) return;

    const size = {
      w: Math.max(300, Math.floor(e.clientX + 5)),
      h: Math.max(400, Math.floor(e.clientY + 5))
    };

    // Send resize message to plugin
    parent.postMessage({
      pluginMessage: {
        type: 'resize',
        size: size
      }
    }, '*');
  };

  const handlePointerDown = (e: PointerEvent) => {
    isResizing = true;
    resizeCorner.setPointerCapture(e.pointerId);

    // Add event listeners for move and up
    resizeCorner.addEventListener('pointermove', resizeWindow);

    // Prevent default to avoid text selection
    e.preventDefault();
  };

  const handlePointerUp = (e: PointerEvent) => {
    if (!isResizing) return;

    isResizing = false;
    resizeCorner.releasePointerCapture(e.pointerId);

    // Remove event listeners
    resizeCorner.removeEventListener('pointermove', resizeWindow);
  };

  // Attach event listeners
  resizeCorner.addEventListener('pointerdown', handlePointerDown);
  resizeCorner.addEventListener('pointerup', handlePointerUp);

  // Handle pointer leave to stop resizing if pointer goes outside
  resizeCorner.addEventListener('pointerleave', (e: PointerEvent) => {
    if (isResizing) {
      handlePointerUp(e);
    }
  });

  console.log('Resize corner functionality initialized');
}

function setupEventListeners() {
  elements.clearSelectionBtn.addEventListener('click', resetTableProperties);
  // elements.reorderColumnsBtn.addEventListener('click', openColumnReorderModal); // Removed - now handled in HTML
  elements.createTableBtn.addEventListener('click', createTable);
  elements.cancelPropsBtn.addEventListener('click', closePropertyEditor);
  elements.propertyEditorOverlay.addEventListener('click', closePropertyEditor);
  elements.savePropsBtn.addEventListener('click', saveCellProperties);
  elements.customCellTextToggle.addEventListener('change', function (this: HTMLInputElement) {
    const show = this.checked;
    elements.customCellTextContainer.style.display = show ? 'block' : 'none';
    // Only show second line if custom cell text is enabled
    if (show) {
      elements.secondLineContainer.style.display = 'flex';
    } else {
      elements.secondLineContainer.style.display = 'none';
      if (elements.secondTextLine) {
        elements.secondTextLine.checked = false;
        elements.secondCellTextContainer.style.display = 'none';
      }
    }
  });
  elements.secondTextLine.addEventListener('change', function (this: HTMLInputElement) {
    elements.secondCellTextContainer.style.display = this.checked ? 'block' : 'none';
  });
  document.querySelectorAll('.apply-option').forEach(option => {
    option.addEventListener('click', handleApplyOptionClick);
  });
  elements.generateSampleCheckbox.addEventListener('change', () => {
    const isChecked = elements.generateSampleCheckbox.checked;
    elements.customCellText.disabled = isChecked;
    elements.aiPromptContainer.style.display = isChecked ? 'block' : 'none';
    // Disable custom cell text toggle when AI generation is enabled
    elements.customCellTextToggle.disabled = isChecked;
  });
  // AI source toggle
  // AI source toggle back
  const aiSource = document.getElementById('aiSource') as HTMLSelectElement | null;
  const fakerConfig = document.getElementById('fakerConfig') as HTMLElement | null;
  const watsonxConfig = document.getElementById('watsonxConfig') as HTMLElement | null;
  aiSource?.addEventListener('change', () => {
    const v = aiSource.value;
    if (fakerConfig) fakerConfig.style.display = v === 'faker' ? 'block' : 'none';
    if (watsonxConfig) watsonxConfig.style.display = v === 'watsonx' ? 'block' : 'none';
  });

  // Load persisted watsonx settings
  parent.postMessage({ pluginMessage: { type: 'load-watsonx-settings' } }, '*');

  // Setup API key synchronization
  setupApiKeySync();

  document.addEventListener('mouseover', function (e) {
    const target = e.target as HTMLElement;
    if (target.classList.contains('cell') || target.classList.contains('header-cell') || target.classList.contains('footer-cell')) {
      showCellTooltip(target as HTMLDivElement);
    }
  });
  document.addEventListener('mouseout', function (e) {
    const target = e.target as HTMLElement;
    if (target.classList.contains('cell') || target.classList.contains('header-cell') || target.classList.contains('footer-cell')) {
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

// Track if user has made any changes that can be reset
let hasChangesToReset = false;

function updateResetButtonState() {
  const resetBtn = elements.clearSelectionBtn;
  if (resetBtn) {
    resetBtn.disabled = !hasChangesToReset;
    resetBtn.style.opacity = hasChangesToReset ? '1' : '0.5';
    resetBtn.style.cursor = hasChangesToReset ? 'pointer' : 'not-allowed';
  }
}

function markChangesForReset() {
  hasChangesToReset = true;
  updateResetButtonState();
}

function clearChangesForReset() {
  hasChangesToReset = false;
  updateResetButtonState();
}

function resetTableProperties() {
  state.cellProperties.clear();
  state.selectedCells.clear();
  state.currentEditingCell = null;
  state.sizeConfirmed = false;

  // If "Populate data from file" is selected, deselect it and clear file selection
  const useFileUpload = document.getElementById('useFileUpload') as HTMLInputElement | null;
  const fileInput = document.getElementById('fileInput') as HTMLInputElement | null;
  const fileUploadContainer = document.getElementById('fileUploadContainer') as HTMLElement | null;
  const filePreview = document.getElementById('filePreview') as HTMLElement | null;

  if (useFileUpload?.checked) {
    useFileUpload.checked = false;
    if (fileUploadContainer) {
      fileUploadContainer.style.display = 'none';
    }
    if (fileInput) {
      fileInput.value = ''; // Clear file selection
    }
    if (filePreview) {
      filePreview.style.display = 'none'; // Hide file preview
    }
  }

  // Recreate the grid to clear all styling including slot colors
  createGrid();
  clearChangesForReset(); // Reset button should be disabled after reset
  showMessage("All cell properties have been reset.", "success");
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

  // Consolidate property collection for ALL cells in the grid
  const propsForFigma: { [key: string]: any } = {};

  // Collect properties for all body cells
  for (let r = 1; r <= state.gridRows; r++) {
    for (let c = 1; c <= state.gridCols; c++) {
      const key = `${r},${c}`;
      const backendKey = `${r - 1}-${c - 1}`;
      const cellData = state.cellProperties.get(key);
      if (cellData) {
        propsForFigma[backendKey] = cellData;
      } else {
        propsForFigma[backendKey] = { properties: {} };
      }
    }
  }

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

  const isUpdate = elements.createTableBtn.textContent === 'Update Table';

  // Only keep scan flow
  const rows = parseInt((document.getElementById('scanRowsInput') as HTMLInputElement)?.value || String(state.gridRows), 10);
  const cols = parseInt((document.getElementById('scanColsInput') as HTMLInputElement)?.value || String(state.gridCols), 10);
  const includeHeader = (document.getElementById('scanHeaderToggle') as HTMLInputElement)?.checked;
  const includeFooter = (document.getElementById('scanFooterToggle') as HTMLInputElement)?.checked;
  const includeSelectable = (document.getElementById('scanSelectableToggle') as HTMLInputElement)?.checked;
  const includeExpandable = (document.getElementById('scanExpandableToggle') as HTMLInputElement)?.checked;

  const message = {
    type: isUpdate ? "update-table" : "create-table-from-scan",
    tableId: isUpdate ? state.tableFrameId : null, // Send table frame ID for updates
    rows,
    cols,
    cellProps: propsForFigma,
    includeHeader,
    includeFooter,
    includeSelectable,
    includeExpandable
  };

  console.log(`🚀 Sending to Figma (${isUpdate ? 'Update' : 'Create'}):`, message);
  console.log(`🚀 cellProps keys:`, Object.keys(propsForFigma));

  parent.postMessage({ pluginMessage: message }, '*');
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
  elements.statusMessage.style.display = 'block'; // Always show when there's a message
}

function hideMessage() {
  elements.statusMessage.style.display = 'none';
}

function showCellTooltip(cell: HTMLDivElement) {
  // Clear any existing timeout
  if (state.hoverTimeout) {
    clearTimeout(state.hoverTimeout);
  }

  // Set a new timeout to show the tooltip after a delay
  state.hoverTimeout = window.setTimeout(() => {
    const row = cell.dataset.row;
    const col = cell.dataset.col;

    let cellKey = '';
    let cellProperties: any = null;

    // Determine the cell key based on the cell type with improved detection
    if (cell.classList.contains('header-cell')) {
      // Handle header cells
      cellKey = `header-${col}`;
      cellProperties = state.cellProperties.get(cellKey);
    } else if (cell.classList.contains('footer-cell')) {
      // Handle footer cells
      cellKey = 'footer';
      cellProperties = state.cellProperties.get(cellKey);
    } else if (row && col) {
      // Handle regular cells
      cellKey = `${row},${col}`;
      cellProperties = state.cellProperties.get(cellKey);
    }

    // If we have cell properties, create tooltip content
    if (cellProperties && cellProperties.properties && Object.keys(cellProperties.properties).length > 0) {
      let tooltipContent = '<div class="tooltip-content">';

      // Add all properties to the tooltip with better formatting
      for (const [key, value] of Object.entries(cellProperties.properties)) {
        // Clean the property name for display
        const cleanKey = key.split('#')[0].trim();

        // Format value based on type
        let displayValue = value;
        if (typeof value === 'object' && value !== null) {
          displayValue = JSON.stringify(value);
        } else if (value === null) {
          displayValue = 'null';
        } else if (value === undefined) {
          displayValue = 'undefined';
        }

        tooltipContent += `
          <div class="tooltip-row">
            <span class="tooltip-key">${cleanKey}</span>
            <span class="tooltip-value">${displayValue}</span>
          </div>
        `;
      }

      tooltipContent += '</div>';

      // Update tooltip content
      state.tooltip.innerHTML = tooltipContent;

      // Get cell position for better tooltip positioning
      const cellRect = cell.getBoundingClientRect();

      // Position the tooltip near the cell
      state.tooltip.style.left = `${cellRect.right + 10}px`;
      state.tooltip.style.top = `${cellRect.top}px`;
      state.tooltip.style.maxWidth = `300px`; // Limit width for better readability

      // Set display to block to show tooltip
      state.tooltip.style.display = 'block';

      // Add tooltip to document if not already added
      if (!state.tooltip.parentElement) {
        document.body.appendChild(state.tooltip);
      }
    } else {
      // If no properties, just hide the tooltip
      hideCellTooltip();
    }
  }, 300); // Reduced delay for better responsiveness
}

function hideCellTooltip() {
  // Clear any existing timeout
  if (state.hoverTimeout) {
    clearTimeout(state.hoverTimeout);
    state.hoverTimeout = null;
  }

  // Hide the tooltip
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
    'Sorted': ['Ascending', 'Descending', 'None'],
    'Sortable': ['True', 'False'],
    'Type': ['Advanced', 'Simple'],
  };

  for (const propName of availableProps) {
    // Skip Size property for all cell types since it can't be updated
    if (propName === 'Size') continue;

    // Skip "Cell text" property since we now have custom cell text functionality
    if (propName.toLowerCase().includes('cell text') && !propName.toLowerCase().includes('second')) {
      continue; // Skip this field entirely
    }

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
      // Skip creating "Show text" checkbox - always hide it
      if (label.toLowerCase().includes('show text')) {
        continue; // Skip this field entirely
      }

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
  // Since we're hiding the "Show text" checkbox, always show the Cell text field
  const cellTextInput = container.querySelector('input[type="text"][id^="dynamic-Cell text"]') as HTMLInputElement;
  if (cellTextInput) {
    const cellTextField = cellTextInput.parentElement as HTMLElement;
    if (cellTextField) {
      cellTextField.style.display = ''; // Always show cell text field
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
  // Since we're hiding the "Show text" checkbox, always show the second text line container
  if (secondTextLineCheckbox && secondTextLineCheckbox.parentElement) {
    const secondLineContainerDyn = secondTextLineCheckbox.parentElement as HTMLElement;
    secondLineContainerDyn.style.display = ''; // Always show second text line container
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

  console.log('[DEBUG] renderDynamicPropertyFieldsFromModel called with props:', props);

  for (const fieldDef of model) {
    // Size property is now available for all apply modes (cell, row, column)
    // if (fieldDef.name === 'Size' && state.applyMode !== 'row') continue;

    const { name, label, type, options, defaultValue, dependsOn, showWhen } = fieldDef;

    // Check if field should be shown based on dependencies
    if (dependsOn && showWhen !== undefined) {
      const dependentValue = props[dependsOn];
      console.log(`[DEBUG] Checking dependency for ${name}: dependsOn=${dependsOn}, showWhen=${showWhen}, currentValue=${dependentValue}`);
      if (String(dependentValue) !== String(showWhen)) {
        console.log(`[DEBUG] Skipping ${name} - dependency condition not met`);
        continue; // Skip this field if dependency condition is not met
      }
      console.log(`[DEBUG] Showing ${name} - dependency condition met`);
    }

    // Use default value if no value is set
    const value = props[name] ?? defaultValue ?? '';

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

      // Add change event listener for all fields to handle dependencies
      select.addEventListener('change', () => {
        // Update the props object with the new value
        props[name] = select.value;
        console.log(`[DEBUG] Field ${name} changed to: ${select.value}`);

        // Special logic: When Sortable is True and Sorted is None, set State to Hover
        if (name === 'Sorted' && select.value === 'None') {
          const sortableInput = document.getElementById('dynamic-Sortable') as HTMLSelectElement;
          if (sortableInput && sortableInput.value === 'True') {
            const stateInput = document.getElementById('dynamic-State') as HTMLSelectElement;
            if (stateInput) {
              stateInput.value = 'Hover';
              props['State'] = 'Hover';
            }
          }
        }

        // Check if this field is a dependency for other fields
        const hasDependentFields = model.some(field => field.dependsOn === name);
        if (hasDependentFields) {
          console.log(`[DEBUG] Field ${name} has dependent fields, re-rendering form`);

          // Collect all current field values before re-rendering
          const currentValues = { ...props };
          for (const fieldDef of model) {
            const input = document.getElementById(`dynamic-${fieldDef.name}`) as HTMLInputElement | HTMLSelectElement;
            if (input) {
              if (fieldDef.type === 'BOOLEAN') {
                currentValues[fieldDef.name] = (input as HTMLInputElement).checked;
              } else {
                currentValues[fieldDef.name] = input.value;
              }
            }
          }

          console.log(`[DEBUG] Collected current values before re-render:`, currentValues);

          // Re-render the form to show/hide dependent fields with preserved values
          renderDynamicPropertyFieldsFromModel(model, currentValues);
        }
      });

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
  console.log('[DEBUG] Final props state:', props);
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
  if (hasShowTextProp && elements.showText) {
    const showTextField = elements.showText.parentElement;
    if (showTextField) showTextField.style.display = 'none';
  }

  // Always show custom cell text toggle for body cells
  if (elements.customCellTextToggle) {
    const customCellTextField = elements.customCellTextToggle.parentElement;
    if (customCellTextField) customCellTextField.style.display = '';
    // Ensure the custom cell text input is shown when toggle is checked
    if (elements.customCellTextToggle.checked && elements.customCellTextContainer) {
      elements.customCellTextContainer.style.display = 'block';
      console.log('[DEBUG] Setting customCellTextContainer display to block in renderBodyCellProperties');
    }
  }

  if (hasCellTextProp) {
    const cellTextField = elements.customCellTextContainer;
    if (cellTextField) cellTextField.style.display = 'none';
  } else {
    // Show cell text container if no dynamic cell text property exists
    const cellTextField = elements.customCellTextContainer;
    if (cellTextField) cellTextField.style.display = 'block';
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
  if (elements.generateSampleCheckbox && elements.generateSampleCheckbox.parentElement) {
    elements.generateSampleCheckbox.parentElement.style.display = '';
  }
  if (elements.aiPromptContainer) {
    elements.aiPromptContainer.style.display = elements.generateSampleCheckbox?.checked ? 'block' : 'none';
  }
  if (elements.colWidthContainer) {
    elements.colWidthContainer.style.display = 'none';
  }
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
  if (elements.showText) {
    const showTextField = elements.showText.parentElement;
    if (showTextField) showTextField.style.display = 'none'; // Always hide show text checkbox
  }
  // Set custom cell text toggle to checked by default and show the input
  if (elements.customCellTextToggle) {
    elements.customCellTextToggle.checked = true;
    // Ensure the input is shown when toggle is checked by default
    if (elements.customCellTextContainer) {
      elements.customCellTextContainer.style.display = cellTextPropKey ? 'none' : 'block';
      console.log('[DEBUG] Setting customCellTextContainer display to', cellTextPropKey ? 'none' : 'block', 'in updateStaticFieldsVisibilityAndValues');
    }
  }
  console.log('[DEBUG] updateStaticFieldsVisibilityAndValues elements presence', {
    customCellText: !!elements.customCellText,
    secondCellText: !!elements.secondCellText,
    secondLineContainer: !!elements.secondLineContainer,
    secondCellTextContainer: !!elements.secondCellTextContainer,
    slotCheckbox: !!elements.slotCheckbox,
    state: !!elements.state
  });
  if (elements.secondLineContainer && elements.secondCellText && elements.customCellTextToggle) {
    elements.secondLineContainer.style.display = secondTextPropKey ? 'none' : (!!elements.secondCellText.value && elements.customCellTextToggle?.checked ? '' : 'none');
  }
  if (elements.secondCellTextContainer && elements.secondCellText && elements.customCellTextToggle) {
    elements.secondCellTextContainer.style.display = secondTextPropKey ? 'none' : (!!elements.secondCellText.value && elements.customCellTextToggle?.checked ? 'block' : 'none');
  }
  const slotField = elements.slotCheckbox?.parentElement;
  if (slotField) slotField.style.display = slotPropKey ? 'none' : '';
  const stateField = elements.state?.parentElement;
  if (stateField) stateField.style.display = statePropKey ? 'none' : '';

  // Populate static fields from props or reset to default
  if (elements.showText) {
    elements.showText.checked = showTextPropKey ? (props[showTextPropKey] || false) : true; // Always default to true since checkbox is hidden
  }
  if (elements.customCellTextToggle) {
    elements.customCellTextToggle.checked = true; // Always default to true for custom cell text
  }
  // Set custom cell text value - show existing cell text if available
  if (elements.customCellText) {
    let existingText = '';
    
    // Try to find existing text from various property keys
    if (cellTextPropKey && props[cellTextPropKey]) {
      existingText = props[cellTextPropKey];
    } else {
      // Look for any text property in the cell properties
      const textProps = Object.keys(props).filter(prop => 
        prop.toLowerCase().includes('text') && 
        !prop.toLowerCase().includes('second') &&
        typeof props[prop] === 'string' &&
        props[prop].trim() !== ''
      );
      if (textProps.length > 0) {
        existingText = props[textProps[0]];
      }
    }
    
    elements.customCellText.value = existingText;
    elements.customCellText.placeholder = 'Enter custom cell text';
  }
  const secondTextValue = secondTextPropKey ? (props[secondTextPropKey] || '') : '';
  if (elements.secondCellText) elements.secondCellText.value = secondTextValue;
  if (elements.secondTextLine) elements.secondTextLine.checked = !!secondTextValue;
  if (elements.slotCheckbox) {
    const slotValue = slotPropKey ? (props[slotPropKey] === true || props[slotPropKey] === 'true') : false;
    elements.slotCheckbox.checked = slotValue;
  }
  if (elements.state) elements.state.value = props['State'] || 'Enabled';

  // Always reset AI-related fields
  if (elements.generateSampleCheckbox) elements.generateSampleCheckbox.checked = false;
  const fakerInput = document.getElementById('fakerMethodInput') as HTMLInputElement | null;
  if (fakerInput) fakerInput.value = '';
  if (elements.aiPromptContainer) elements.aiPromptContainer.style.display = 'none';
  if (elements.customCellText) elements.customCellText.disabled = false;
  if (elements.customCellTextToggle) elements.customCellTextToggle.disabled = false;

  // --- For static property fields ---
  if (elements.secondLineContainer && elements.secondCellText && elements.secondCellTextContainer && elements.secondTextLine && !secondTextPropKey) {
    elements.secondLineContainer.style.display = !!elements.secondCellText.value ? '' : 'none';
    if (!elements.secondCellText.value) {
      elements.secondTextLine.checked = false;
      elements.secondCellTextContainer.style.display = 'none';
    }
  }
}

function openPropertyEditor(key: string) {
  if (state.mode !== 'edit') return;

  // If we don't have component info and this is a body cell, request it
  if (!state.selectedComponent && !key.startsWith('header-') && key !== 'footer') {
    console.log(`[openPropertyEditor] No component info available, requesting...`);
    parent.postMessage({
      pluginMessage: {
        type: 'request-component-info',
        tableId: state.tableFrameId
      }
    }, '*');

    // Wait for component info to be received before proceeding (with timeout)
    let attempts = 0;
    const maxAttempts = 20; // 1 second timeout (20 * 50ms)

    const waitForComponentInfo = () => {
      attempts++;
      if (state.selectedComponent) {
        console.log(`[openPropertyEditor] Component info received, proceeding with editor`);
        openPropertyEditorInternal(key); // Call the internal function now that we have component info
      } else if (attempts < maxAttempts) {
        setTimeout(waitForComponentInfo, 50);
      } else {
        console.log(`[openPropertyEditor] Timeout waiting for component info, proceeding with fallback`);
        // Proceed with the editor anyway, using fallback logic
        openPropertyEditorInternal(key);
      }
    };
    waitForComponentInfo();
    return;
  }

  // Call the internal function to avoid infinite recursion
  openPropertyEditorInternal(key);
}

function openPropertyEditorInternal(key: string) {
  try {
    // Guarantee the editor has required nodes
    ensurePropertyEditorScaffold();
    // Show column width option based on current apply mode
    if (state.applyMode === 'cell' || state.applyMode === 'column') {
      elements.colWidthContainer.style.display = 'block';
      console.log(`[DEBUG] Initial: Showing column width option for ${state.applyMode} mode`);
    } else {
      elements.colWidthContainer.style.display = 'none';
      console.log(`[DEBUG] Initial: Hiding column width option for ${state.applyMode} mode`);
    }
    // Set default apply mode and update tab visual state
    if (key.startsWith('header-') || key === 'footer') {
      // For header/footer cells, default to 'cell' mode
      state.applyMode = 'cell';
      // Update tab visual state
      document.querySelectorAll('.apply-option').forEach(opt => opt.classList.remove('active'));
      const cellOption = document.querySelector('.apply-option[data-apply="cell"]') as HTMLElement;
      if (cellOption) {
        cellOption.classList.add('active');
      }
    } else {
      // For body cells, default to 'column' mode
      state.applyMode = 'column';
      // Update tab visual state
      document.querySelectorAll('.apply-option').forEach(opt => opt.classList.remove('active'));
      const columnOption = document.querySelector('.apply-option[data-apply="column"]') as HTMLElement;
      if (columnOption) {
        columnOption.classList.add('active');
      }
    }

    // Show/hide apply options based on cell type
    const applyOptionsContainer = document.querySelector('.apply-options') as HTMLElement;
    if (key === 'footer' || key.startsWith('header-')) {
      // Hide apply options for footer and header cells
      if (applyOptionsContainer) {
        applyOptionsContainer.style.display = 'none';
      }
    } else {
      // Show apply options for body cells
      if (applyOptionsContainer) {
        applyOptionsContainer.style.display = '';
      }
    }

    // Show all apply options for body cells, hide column option for header cells
    const applyToCellOption = document.querySelector('.apply-option[data-apply="cell"]') as HTMLElement;
    const applyToRowOption = document.querySelector('.apply-option[data-apply="row"]') as HTMLElement;
    const applyToColumnOption = document.querySelector('.apply-option[data-apply="column"]') as HTMLElement;

    if (key.startsWith('header-')) {
      // For header cells, hide all apply options (we already hide the container above)
      if (applyToCellOption) applyToCellOption.style.display = 'none';
      if (applyToRowOption) applyToRowOption.style.display = 'none';
      if (applyToColumnOption) applyToColumnOption.style.display = 'none';
    } else {
      // For body cells, show all options
      // For footer and header, we've already hidden the entire container, but let's make sure
      if (applyToCellOption) applyToCellOption.style.display = (key === 'footer' || key.startsWith('header-')) ? 'none' : '';
      if (applyToRowOption) applyToRowOption.style.display = (key === 'footer' || key.startsWith('header-')) ? 'none' : '';
      if (applyToColumnOption) applyToColumnOption.style.display = (key === 'footer' || key.startsWith('header-')) ? 'none' : '';
    }

    console.log(`[DEBUG] Apply options visibility for ${key}:`, {
      cell: applyToCellOption?.style.display,
      row: applyToRowOption?.style.display,
      column: applyToColumnOption?.style.display
    });

    // Re-attach event listeners to apply options in case they were recreated
    const applyOptions = document.querySelectorAll('.apply-option');
    console.log(`[DEBUG] Found ${applyOptions.length} apply options to attach listeners to`);

    applyOptions.forEach((option, index) => {
      console.log(`[DEBUG] Attaching listener to option ${index}:`, option.textContent);
      // Remove existing listeners to avoid duplicates
      option.removeEventListener('click', handleApplyOptionClick);
      // Add new listener
      option.addEventListener('click', handleApplyOptionClick);
      console.log(`[DEBUG] Listener attached to:`, option.textContent);
    });

    state.currentEditingCell = key;

    let availableProps: string[] = state.selectedComponent?.availableProperties || [];
    let propertyTypes: { [key: string]: any } = state.selectedComponent?.propertyTypes || {};

    console.log(`[DEBUG] Available properties for ${key}:`, availableProps);
    console.log(`[DEBUG] Property types for ${key}:`, propertyTypes);
    console.log(`[DEBUG] Size property available:`, availableProps.includes('Size'));

    // Add Size property if not available (for testing purposes)
    if (!availableProps.includes('Size')) {
      console.log(`[DEBUG] Adding Size property as fallback`);
      availableProps.push('Size');
      propertyTypes['Size'] = 'VARIANT';
    }

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

      // Apply default values for header cell properties
      for (const fieldDef of HEADER_CELL_MODEL) {
        if (fieldDef.defaultValue && props[fieldDef.name] === undefined) {
          props[fieldDef.name] = fieldDef.defaultValue;
        }
      }

      label = `Header ${key.split('-')[1]}`;
      console.log('[DEBUG] openPropertyEditor header', { props });
      renderDynamicPropertyFieldsFromModel(HEADER_CELL_MODEL, props);
    } else if (key === 'footer') {
      const cellState = state.cellProperties.get(key);
      props = (cellState && cellState.properties) ? cellState.properties : {};
      label = 'Footer';
      console.log('[DEBUG] openPropertyEditor footer', { props });

      // For footer, only show the "Type" property
      const footerModel = FOOTER_MODEL.filter(field => field.label === "Type");
      renderDynamicPropertyFieldsFromModel(footerModel, props);
    } else {
      const cellState = getCellState(key);
      props = cellState.properties || {};
      const [row, col] = key.split(',');
      label = `(${row},${col})`;
      console.log('[DEBUG] openPropertyEditor body', { availableProps, propertyTypes, props });
      renderDynamicPropertyFields(availableProps, propertyTypes, props);
    }

    // Show col width for all cells (body, header, footer)
    let width: number | undefined = undefined;

    if (key.startsWith('header-')) {
      // For header cells, get width from the corresponding column
      const col = key.split('-')[1];
      const colCells = Array.from({ length: state.gridRows }, (_, r) => `${r + 1},${col}`);
      for (const k of colCells) {
        const cellState = state.cellProperties.get(k);
        if (cellState && cellState.colWidth) {
          width = cellState.colWidth;
          break;
        }
      }
      if (width === undefined) {
        width = 120;
      }
      console.log('[DEBUG] Header col width UI exists?', !!elements.colWidthContainer, !!elements.colWidthInput, 'value to set:', width);
      if (elements.colWidthContainer) elements.colWidthContainer.style.display = 'block';
      if (elements.colWidthInput) elements.colWidthInput.value = String(width);
      console.log(`[DEBUG] Header cell ${key} - colWidth: ${width}, container display: ${elements.colWidthContainer.style.display}`);
    } else if (key === 'footer') {
      // For footer, use a default width
      width = 120;
      console.log('[DEBUG] Footer col width UI exists?', !!elements.colWidthContainer, !!elements.colWidthInput, 'value to set:', width);
      // Hide column width for footer cells
      if (elements.colWidthContainer) elements.colWidthContainer.style.display = 'none';
      if (elements.colWidthInput) elements.colWidthInput.value = String(width);
      console.log(`[DEBUG] Footer cell - colWidth: ${width}, container display: ${elements.colWidthContainer.style.display}`);
    } else {
      // For body cells, get width from the column
      const col = key.split(',')[1];
      const colCells = Array.from({ length: state.gridRows }, (_, r) => `${r + 1},${col}`);
      for (const k of colCells) {
        const cellState = state.cellProperties.get(k);
        if (cellState && cellState.colWidth) {
          width = cellState.colWidth;
          break;
        }
      }
      if (width === undefined) {
        width = 120;
      }
      console.log('[DEBUG] Body col width UI exists?', !!elements.colWidthContainer, !!elements.colWidthInput, 'value to set:', width, 'for key', key);
      if (elements.colWidthContainer) elements.colWidthContainer.style.display = 'none';
      if (elements.colWidthInput) elements.colWidthInput.value = String(width);
      console.log(`[DEBUG] Body cell ${key} - colWidth: ${width}, container display: ${elements.colWidthContainer.style.display}`);
    }

    // Set title and show editor
    elements.propertyEditorTitle.innerHTML = `Cell Properties <span>${label}</span>`;
    elements.editingCellCoords.textContent = label;
    elements.propertyEditorOverlay.style.display = 'block';
    elements.propertyEditor.style.display = 'block';
    // Single-prompt UI wiring moved to main page; no wiring here

    // Final check: Ensure custom cell text input is shown for body cells if toggle is checked
    if (!key.startsWith('header-') && key !== 'footer') {
      if (elements.customCellTextToggle && elements.customCellTextToggle.checked && elements.customCellTextContainer) {
        elements.customCellTextContainer.style.display = 'block';
        console.log('[DEBUG] Final check: Setting customCellTextContainer display to block for body cell');
      }
    }

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
      const cellState = getCellState(key);
      const existingProps = cellState.properties || {};
      const newProps: any = { ...existingProps };
      
      for (const fieldDef of HEADER_CELL_MODEL) {
        const { name, type, defaultValue } = fieldDef;
        const input = document.getElementById(`dynamic-${name}`) as HTMLInputElement | HTMLSelectElement;
        if (!input) continue;
        if (type === 'BOOLEAN') {
          newProps[name] = (input as HTMLInputElement).checked;
        } else {
          const inputValue = input.value;
          if (inputValue || !existingProps[name]) {
            newProps[name] = inputValue || defaultValue || '';
          }
        }
      }

      // Handle column width for header cells
      let colWidth: number | undefined = undefined;
      if (elements.colWidthInput && elements.colWidthInput.value) {
        colWidth = parseInt(elements.colWidthInput.value, 10);
      } else {
        console.log('[DEBUG] colWidthInput missing or empty when saving cell (footer)');
      }

      cellState.properties = newProps;

      // If column width is set, apply it to the entire column
      if (colWidth) {
        const col = key.split('-')[1];
        const colCells = Array.from({ length: state.gridRows }, (_, r) => `${r + 1},${col}`);
        for (const k of colCells) {
          const colCellState = getCellState(k);
          colCellState.colWidth = colWidth;
        }
        // Also set it for the header cell itself
        cellState.colWidth = colWidth;
      }

      console.log('[DEBUG] saveCellProperties header', key, newProps, 'colWidth:', colWidth);
      state.cellProperties.set(key, cellState);
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

      // For footer cells, we don't process column width since the input is hidden
      // Just save the properties without any column width changes
      const cellState = getCellState(key);
      cellState.properties = newProps;

      console.log('[DEBUG] saveCellProperties footer', key, newProps);
      state.cellProperties.set(key, cellState);

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

      // Save custom cell text if provided (regardless of dynamic properties)
      const showText = elements.customCellTextToggle.checked;
      const generateSampleData = elements.generateSampleCheckbox.checked;

      // Find the Show text property
      const showTextProp = availableProps.find(p => {
        const type = propertyTypes[p];
        return type === 'BOOLEAN' && p.toLowerCase().includes('show') && p.toLowerCase().includes('text');
      });

      // Turn off Show text if both custom text and AI generation are off
      if (showTextProp && !showText && !generateSampleData) {
        props[showTextProp] = false;
        console.log('[DEBUG] Turning off Show text property:', showTextProp);
      } else if (showTextProp && (showText || generateSampleData)) {
        props[showTextProp] = true;
        console.log('[DEBUG] Turning on Show text property:', showTextProp);
      }

      if (showText && elements.customCellText.value) {
        // Find the correct property key for cell text
        let cellTextProp = availableProps.find(p => {
          const type = propertyTypes[p];
          return type === 'TEXT' && p.toLowerCase().includes('text') && !p.toLowerCase().includes('second');
        }) || availableProps.find(p => p.toLowerCase().includes('text')) || 'Cell text#12234:16';
        props[cellTextProp] = elements.customCellText.value;
        console.log('[DEBUG] Saving custom cell text:', elements.customCellText.value, 'to property:', cellTextProp);
      }
      if (showText && elements.secondTextLine.checked && elements.secondCellText.value) {
        // Find the correct property key for second text
        let secondTextProp = availableProps.find(p => p.toLowerCase().includes('second')) || 'secondCellText';
        props[secondTextProp] = elements.secondCellText.value;
      }
      if (!hasStateProp) {
        props['State'] = elements.state.value;
      }

      // Save static slot property if not handled by dynamic properties
      if (!hasSlotProp && elements.slotCheckbox) {
        const slotPropName = availableProps.find(p => p.toLowerCase().includes('slot')) || 'Slot';
        props[slotPropName] = elements.slotCheckbox.checked;
        console.log(`🔍 UI: Saving static Slot property: ${slotPropName} = ${props[slotPropName]}`);
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

        // Special logging for Slot properties
        if (propName.toLowerCase().includes('slot')) {
          console.log(`🔍 UI: Saving Slot property: ${propName} = ${props[propName]} (type: ${type})`);
        }
      }

      // Col width logic: always set colWidth as a top-level property on first row of column
      let colWidth: number | undefined = undefined;
      if (elements.colWidthInput && elements.colWidthInput.value) {
        colWidth = parseInt(elements.colWidthInput.value, 10);
      } else {
        console.log('[DEBUG] colWidthInput missing or empty when saving cell (body)');
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
          const k = `${row},${c}`;
          if (state.selectedCells.has(k)) {
            if (colWidth && row === 1) {
              applyProps(k, { ...props }, colWidth);
            } else {
              applyProps(k, { ...props });
            }
            const cellState = getCellState(k);
            if (cellState.isCheckbox !== undefined) {
              cellState.isCheckbox = cellState.isCheckbox;
            }
          }
        }
      } else if (state.applyMode === 'column') {
        for (let r = 1; r <= state.gridRows; r++) {
          const k = `${r},${col}`;
          if (state.selectedCells.has(k)) {
            if (colWidth && r === 1) {
              applyProps(k, { ...props }, colWidth);
            } else {
              applyProps(k, { ...props });
            }
            const cellState = getCellState(k);
            if (cellState.isCheckbox !== undefined) {
              cellState.isCheckbox = cellState.isCheckbox;
            }
          }
        }
      }

      // Now handle AI sample data generation
      const generateAI = elements.generateSampleCheckbox.checked;
      const aiChoice = (document.getElementById('aiSource') as HTMLSelectElement | null)?.value || 'faker';
      const fakerMethod = fakerMethodInput.value.trim();
      if (generateAI && aiChoice === 'faker' && fakerMethod) {
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
      if (generateAI && aiChoice === 'watsonx') {
        let count = 1;
        let mode: 'cell' | 'row' | 'column' = 'cell';
        if (state.applyMode === 'row') {
          count = state.gridCols;
          mode = 'row';
        } else if (state.applyMode === 'column') {
          count = state.gridRows;
          mode = 'column';
        }
        const prompt = (document.getElementById('watsonxPrompt') as HTMLTextAreaElement | null)?.value || 'Generate short realistic values';
        const endpoint = 'https://us-south.ml.cloud.ibm.com';
        const apiKey = (document.getElementById('watsonxApiKey') as HTMLInputElement | null)?.value || '';
        const useAccessToken = false;
        const accessToken = '';
        if (!apiKey) {
          showMessage('Provide IBM Cloud API key', 'error');
          return;
        }
        const remember = true;
        const useProxy = true;
        const proxyUrl = 'http://localhost:3000';
        pendingFakerContext = { mode, key, fakerMethod: 'watsonx' };
        parent.postMessage({ pluginMessage: { type: 'generate-watsonx-data', prompt, endpoint, apiKey, accessToken, useAccessToken, useProxy, proxyUrl, count, remember } }, '*');
        return;
      }
    }
    closePropertyEditor();
    updateCellVisuals();
    markChangesForReset(); // Enable reset button after cell properties are saved
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

  // Refresh the grid to show visual changes (like slot styling)
  createGrid();
  
  // Refresh component info to ensure preview grid shows updated values
  if (state.tableFrameId) {
    parent.postMessage({
      pluginMessage: {
        type: 'request-component-info',
        tableId: state.tableFrameId
      }
    }, '*');
  }
}

function updateCellVisuals() {
  const availableProps: string[] = state.selectedComponent?.availableProperties || [];
  console.log(`[updateCellVisuals] availableProps:`, availableProps);
  console.log(`[updateCellVisuals] selectedComponent:`, state.selectedComponent);

  // Update body cells
  document.querySelectorAll('.cell').forEach(cell => {
    const divCell = cell as HTMLDivElement;
    const key = `${divCell.dataset.row},${divCell.dataset.col}`;
    const props = state.cellProperties.get(key);
    console.log(`[updateCellVisuals] Cell ${key}:`, props);

    if (props && props.properties && Object.keys(props.properties).length > 0) {
      let displayText = '';
      let hasSlotEnabled = false;

      // If we have selectedComponent, use its property types
      if (state.selectedComponent && availableProps.length > 0) {
        // Check for slot property first
        const slotProp = availableProps.find(p =>
          state.selectedComponent!.propertyTypes[p] === 'BOOLEAN' && p.toLowerCase().includes('slot'));
        if (slotProp && (props.properties[slotProp] === true || props.properties[slotProp] === 'true')) {
          hasSlotEnabled = true;
        }

        // Find any TEXT property to display as cell text
        for (const propName of availableProps) {
          const type = state.selectedComponent.propertyTypes[propName];
          if (type === 'TEXT' && props.properties[propName]) {
            displayText += props.properties[propName];
            break; // Use the first TEXT property found
          }
        }
        // If no text and slot is checked, show empty
        if (!displayText && hasSlotEnabled) {
          displayText = '';
        }
      } else {
        // Fallback: look for common text property names in the loaded properties
        const textProps = Object.keys(props.properties).filter(prop =>
          prop.toLowerCase().includes('text') &&
          !prop.toLowerCase().includes('second') &&
          props.properties[prop] &&
          typeof props.properties[prop] === 'string' &&
          props.properties[prop].trim() !== ''
        );

        if (textProps.length > 0) {
          displayText = props.properties[textProps[0]];
        }

        // Check for slot property
        const slotProps = Object.keys(props.properties).filter(prop =>
          prop.toLowerCase().includes('slot')
        );
        if (slotProps.length > 0 && (props.properties[slotProps[0]] === true || props.properties[slotProps[0]] === 'true')) {
          hasSlotEnabled = true;
          if (!displayText) {
            displayText = '';
          }
        }
      }

      if (hasSlotEnabled) {
        divCell.classList.add('slot-enabled');
        divCell.classList.remove('edited');
        if (!displayText) {
          displayText = ''; // Show empty for slot-enabled cells
        }
      } else if (!displayText) {
        displayText = `${divCell.dataset.row},${divCell.dataset.col}`;
        divCell.classList.remove('edited');
        divCell.classList.remove('slot-enabled');
        divCell.style.fontWeight = 'normal';
      } else {
        divCell.classList.add('edited');
        divCell.classList.remove('slot-enabled');
        divCell.style.fontWeight = 'bold';
      }
      divCell.textContent = displayText;
    } else {
      divCell.textContent = `${divCell.dataset.row},${divCell.dataset.col}`;
      divCell.classList.remove('edited');
      divCell.classList.remove('slot-enabled');
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
        let displayText = '';

        // Try to get text from properties
        if (props && props.properties) {
          // Use header cell component properties if available
          let textKey = '';
          if (state.headerCellComponent?.availableProperties && state.headerCellComponent?.propertyTypes) {
            const availableProps: string[] = state.headerCellComponent.availableProperties;
            const propertyTypes: { [key: string]: any } = state.headerCellComponent.propertyTypes;
            textKey = availableProps.find(p => propertyTypes[p] === 'TEXT') || '';
          }

          // If we couldn't find the text key from the component, try to find it in the properties
          if (!textKey) {
            textKey = Object.keys(props.properties).find(p =>
              p.toLowerCase().includes('text') &&
              !p.toLowerCase().includes('second') &&
              typeof props.properties[p] === 'string'
            ) || '';
          }

          // Try multiple fallbacks to find the text
          if (textKey && props.properties[textKey]) {
            displayText = props.properties[textKey];
          } else if (props.properties['Cell text#12234:32']) {
            // Directly check for the hardcoded property name
            displayText = props.properties['Cell text#12234:32'];
          } else {
            // Check for any property that might contain text
            const textProps = Object.keys(props.properties).filter(p =>
              typeof props.properties[p] === 'string' &&
              props.properties[p].toString().trim() !== ''
            );

            if (textProps.length > 0) {
              displayText = props.properties[textProps[0]].toString();
            }
          }
        }

        // Set the display text or fallback to default
        if (displayText && typeof displayText === 'string' && displayText.trim().length > 0) {
          // Decode HTML entities to prevent &amp;quot; issues
          const decodedText = displayText.replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>');
          cell.textContent = decodedText;
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



// Message handling
window.onmessage = (event) => {
  if (!domReady || !elements.grid) return;
  const msg = event.data.pluginMessage;
  if (!msg) return;

  console.log(`[UI] Received message: ${msg.type}`, msg);

  switch (msg.type) {
    case 'ai-table-response': {
      hideLoader();
      const headers: string[] = (msg.headers || []).map(String);
      const rowsData: string[][] = Array.isArray(msg.rows) ? msg.rows.map((r: any) => Array.isArray(r) ? r.map(String) : []) : [];
      if (!headers.length || !rowsData.length) {
        showMessage('AI did not return usable table data.', 'error');
        break;
      }
      // Helper to sanitize noisy AI strings
      const sanitize = (s: string): string => {
        if (!s) return '';
        let t = String(s).trim();
        t = t.replace(/^```[a-zA-Z]*\n?|```$/g, '').trim();
        if (t.includes(':')) {
          const parts = t.split(':');
          t = parts[parts.length - 1];
        }
        t = t.replace(/[\[\]\{\}]/g, '').replace(/^"|"$/g, '').replace(/^'|'$/g, '');
        t = t.trim().replace(/^"|"$/g, '').replace(/^'|'$/g, '');
        t = t.replace(/\s+/g, ' ').trim();
        return t;
      };

      // Enforce the selected grid size; do not resize to AI result
      const desiredRows = state.gridRows;
      const desiredCols = state.gridCols;

      // Build fixed headers sized exactly to desiredCols
      const fixedHeaders: string[] = [];
      for (let c = 0; c < desiredCols; c++) {
        fixedHeaders.push(sanitize(headers[c] || `Column ${c + 1}`));
      }

      // Build fixed rows sized exactly desiredRows x desiredCols
      const fixedRows: string[][] = [];
      for (let r = 0; r < desiredRows; r++) {
        const src = rowsData[r] || [];
        const row: string[] = [];
        for (let c = 0; c < desiredCols; c++) {
          row.push(sanitize(src[c] || ''));
        }
        fixedRows.push(row);
      }

      // Render into current grid
      createGrid();
      elements.gridContainer.style.display = 'flex';
      elements.actionButtons.style.display = 'flex';

      // Use header cell component properties if available, otherwise fallback to body cell properties
      const headerAvailableProps: string[] = state.headerCellComponent?.availableProperties || state.selectedComponent?.availableProperties || [];
      const headerPropertyTypes: { [key: string]: any } = state.headerCellComponent?.propertyTypes || state.selectedComponent?.propertyTypes || {};
      const bodyAvailableProps: string[] = state.selectedComponent?.availableProperties || [];
      const bodyPropertyTypes: { [key: string]: any } = state.selectedComponent?.propertyTypes || {};

      const headerTextKey = headerAvailableProps.find(p => headerPropertyTypes[p] === 'TEXT') || 'Cell text';
      const bodyTextKey = bodyAvailableProps.find(p => bodyPropertyTypes[p] === 'TEXT') || 'Cell text';

      // Apply header cell properties
      for (let c = 1; c <= desiredCols; c++) {
        const key = `header-${c}`;
        const cellState = getCellState(key);
        cellState.properties[headerTextKey] = fixedHeaders[c - 1] || `H${c}`;
      }

      // Apply body cell properties
      for (let r = 0; r < desiredRows; r++) {
        for (let c = 0; c < desiredCols; c++) {
          const key = `${r + 1},${c + 1}`;
          const cellState = getCellState(key);
          cellState.properties[bodyTextKey] = fixedRows[r][c] || '';
        }
      }

      updateCellVisuals();
      if (elements.createTableBtn) elements.createTableBtn.disabled = false;
      markChangesForReset(); // Enable reset button after AI content is applied
      showMessage(`AI content applied (clamped to ${desiredRows} rows x ${desiredCols} cols). Review/edit then click Create Table.`, 'success');
      break;
    }
    case "table-selected":
      // Handle when a Data table component is selected
      state.hasComponent = true;
      showMessage(`Selected: Data table`, "success");
      elements.gridContainer.style.display = "none";
      elements.actionButtons.style.display = "none";
      // Trigger scan-table to analyze the selected table with a small delay
      showLoader('Scanning table...');
      setTimeout(() => {
        parent.postMessage({ pluginMessage: { type: 'scan-table' } }, '*');
      }, 100);
      break;



    case "component-selected":
      // Reset state when a new component is selected
      state.cellProperties.clear();
      state.selectedCells.clear();
      state.currentEditingCell = null;
      state.sizeConfirmed = false;
      state.tableFrameId = undefined;
      clearChangesForReset(); // Disable reset button for new component
      console.log('[DEBUG] Reset state for new component selection');

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
      // Only show error if we're not currently editing an existing table
      if (!state.tableFrameId) {
        showMessage("Please select a table cell component.", "error");
        elements.gridContainer.style.display = "none";
        elements.actionButtons.style.display = "none";
        state.selectedComponent = null;
        elements.landingPage.style.display = 'block';
        hideMessage(); // Hide status messages on landing page
      }
      break;

    case "table-created":
      hideLoader();
      if (msg.isComponent) {
        showMessage("Table component created successfully! You can now reuse it.", "success");
      } else {
        showMessage("Table created successfully!", "success");
      }
      // Show landing page and componentModeBtn again for new table generation
      elements.landingPage.style.display = 'flex';
      hideMessage(); // Hide status messages on landing page
      elements.componentModeBtn.style.display = 'inline-block';
      elements.gridContainer.style.display = 'none';
      elements.actionButtons.style.display = 'none';
      elements.propertyEditor.style.display = 'none';
      // Clear all cell properties after table is generated
      state.cellProperties.clear();
      // Reset button text to "Create Table" for new tables
      elements.createTableBtn.textContent = 'Create Table';
      state.tableFrameId = undefined;
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
            // Always apply AI data, clearing any existing text
            // This ensures the latest AI-generated data takes precedence
            cellState.properties[textProp] = data;
            console.log(`🔄 Applied AI data to cell ${cellKey}: "${data}"`);
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
            applyAiData(k, sample[(c - 1) % sample.length]);
          }
        } else if (mode === 'column') {
          const [, col] = key.split(',').map(Number);
          for (let r = 1; r <= state.gridRows; r++) {
            const k = `${r},${col}`;
            applyAiData(k, sample[(r - 1) % sample.length]);
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

    case "watsonx-data-response":
      hideLoader();
      const wx = msg.data;
      if (!Array.isArray(wx) || wx.length === 0) {
        showMessage('No data generated from watsonx.ai.', 'error');
        break;
      }
      if (pendingFakerContext) {
        const { mode, key } = pendingFakerContext;
        const availableProps: string[] = state.selectedComponent?.availableProperties || [];
        const propertyTypes: { [key: string]: any } = state.selectedComponent?.propertyTypes || {};
        function getCellTextProp() { return availableProps.find(p => propertyTypes[p] === 'TEXT'); }
        function getTextVisibilityProp() { return availableProps.find(p => propertyTypes[p] === 'BOOLEAN' && (p.toLowerCase().includes('show') || p.toLowerCase().includes('text')) && !p.toLowerCase().includes('slot')); }
        const applyAiData = (cellKey: string, data: string) => {
          const cellState = getCellState(cellKey);
          const textProp = getCellTextProp();
          const visibilityProp = getTextVisibilityProp();
          if (textProp) cellState.properties[textProp] = data;
          if (visibilityProp) cellState.properties[visibilityProp] = true;
          state.selectedCells.add(cellKey);
        };
        if (mode === 'cell') {
          applyAiData(key, wx[0]);
        } else if (mode === 'row') {
          const [row] = key.split(',').map(Number);
          for (let c = 1; c <= state.gridCols; c++) {
            const k = `${row},${c}`;
            applyAiData(k, wx[(c - 1) % wx.length]);
          }
        } else if (mode === 'column') {
          const [, col] = key.split(',').map(Number);
          for (let r = 1; r <= state.gridRows; r++) {
            const k = `${r},${col}`;
            applyAiData(k, wx[(r - 1) % wx.length]);
          }
        }
        pendingFakerContext = null;
        updateCellVisuals();
        closePropertyEditor();
        document.querySelectorAll('.cell').forEach(cell => {
          const divCell = cell as HTMLDivElement;
          const k = `${divCell.dataset.row},${divCell.dataset.col}`;
          if (state.selectedCells.has(k)) divCell.classList.add('selected');
        });
      }
      break;

    case "prompt-fallback-notice":
      hideLoader();
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

      // Hide landing page and show status message
      elements.landingPage.style.display = 'none';
      elements.statusMessage.style.display = 'block';
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

      // Reset state for new table generation
      state.cellProperties.clear();
      state.selectedCells.clear();
      state.currentEditingCell = null;
      state.sizeConfirmed = false;
      state.tableFrameId = undefined;
      console.log('[DEBUG] Reset state for new table generation');

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

        // Update the grid to reflect any changes in component properties
        renderHeaderFooterGrids();
      } else {
        showMessage('Could not find a body cell template in the scanned table.', 'error');
        if (elements.scanTableSection) elements.scanTableSection.style.display = 'flex';
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
              <label style="margin-right: 8px;">Rows: <input type="number" id="scanRowsInput" class="styled-input" value="${state.gridRows}" min="1" style="width: 70px;"></label>
              <label>Cols: <input type="number" id="scanColsInput" class="styled-input" value="${state.gridCols}" min="1" style="width: 70px;"></label>
          </div>
          <div style="display: flex; align-items: center; gap: 16px; flex-wrap: wrap; margin-top: 8px; margin-bottom: 8px;">
              <div class="property-field checkbox">
                <input type="checkbox" id="scanHeaderToggle" checked>
                <label for="scanHeaderToggle">Header</label>
              </div>
              <div class="property-field checkbox">
                <input type="checkbox" id="scanFooterToggle" ${details.footer ? 'checked' : ''}>
                <label for="scanFooterToggle">Footer</label>
              </div>
              <div class="property-field checkbox">
                <input type="checkbox" id="scanSelectableToggle">
                <label for="scanSelectableToggle">Selectable</label>
              </div>
              <div class="property-field checkbox">
                <input type="checkbox" id="scanExpandableToggle">
                <label for="scanExpandableToggle">Expandable</label>
              </div>
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

      // Change button text to "Create" and disable update
      elements.createTableBtn.textContent = 'Create Table';
      elements.createTableBtn.disabled = false;
      // Optionally, you could hide the button if you have two separate buttons
      // elements.createTableBtn.style.display = 'inline-block';

      setMode('edit');
      updateCreateButtonState();
      break;

    case "generated-table-selected-no-settings":
      console.log(`[UI] Generated table selected but no settings available`);
      state.tableFrameId = msg.tableId;
      state.hasComponent = true;

      // Show UI without error message
      elements.landingPage.style.display = 'none';
      elements.gridContainer.style.display = 'flex';
      elements.actionButtons.style.display = 'flex';

      // Reset to default state
      state.gridRows = 5;
      state.gridCols = 5;
      state.cellProperties.clear();
      createGrid();

      // Change button text to indicate this is an update
      elements.createTableBtn.textContent = "Update Table";
      elements.createTableBtn.disabled = false;

      showMessage("Generated table selected. You can modify and update it.", "success");
      break;

    case "edit-existing-table":
      console.log(`[UI] Processing edit-existing-table with settings:`, msg.settings);
      console.log(`[UI] cellProperties keys:`, Object.keys(msg.settings.cellProperties || {}));
      const settings = msg.settings;
      if (!settings) return;
      state.tableFrameId = msg.tableId;

      // Reset state
      state.selectedCells.clear();
      state.sizeConfirmed = false;
      state.currentEditingCell = null;
      state.hasComponent = true; // Ensure this is true for generated tables

      // Show UI
      elements.landingPage.style.display = 'none';
      elements.gridContainer.style.display = 'flex';
      elements.actionButtons.style.display = 'flex';

      console.log(`[UI] Showing grid and action buttons. Grid display: ${elements.gridContainer.style.display}, Action buttons display: ${elements.actionButtons.style.display}`);

      // Update state from settings
      state.gridRows = settings.rows || 5;
      state.gridCols = settings.columns || 5;

      // Convert backend cell property keys (0-0, 0-1) to UI format (1,1, 1,2)
      const convertedCellProperties = new Map();
      if (settings.cellProperties) {
        Object.entries(settings.cellProperties).forEach(([key, value]) => {
          if (key === 'footer') {
            // Keep footer as is
            convertedCellProperties.set(key, value);
          } else if (key.startsWith('header-')) {
            // Keep header keys as is
            convertedCellProperties.set(key, value);
          } else if (key.includes('-')) {
            // Convert "0-0" to "1,1" format (only for body cells)
            const [row, col] = key.split('-').map(Number);
            const uiKey = `${row + 1},${col + 1}`;
            convertedCellProperties.set(uiKey, value);
          }
        });
      }
      state.cellProperties = convertedCellProperties;
      console.log(`[UI] Converted cell properties:`, Array.from(convertedCellProperties.keys()));

      // Create grid and update visuals
      createGrid();

      // Request component information from backend for proper display
      console.log(`[UI] Requesting component info from backend...`);
      parent.postMessage({
        pluginMessage: {
          type: 'request-component-info',
          tableId: state.tableFrameId
        }
      }, '*');

      // Update visuals after a short delay to allow component info to load
      setTimeout(() => {
        updateCellVisuals();
      }, 100);

      // Create scan options for editing existing table
      if (!elements.scanOptionsContainer) {
        const optionsDiv = document.createElement('div');
        optionsDiv.id = 'scanOptionsContainer';
        optionsDiv.style.display = 'flex';
        optionsDiv.style.flexDirection = 'column';
        optionsDiv.style.alignItems = 'flex-start';
        optionsDiv.style.marginBottom = '12px';
        optionsDiv.style.gap = '0px';
        optionsDiv.innerHTML = `
            <div style="display: flex; align-items: center; gap: 10px; flex-wrap: wrap; margin-bottom: 8px; width: 100%; justify-content: flex-start;">
                <label style="margin-right: 8px;">Rows: <input type="number" id="scanRowsInput" class="styled-input" value="${state.gridRows}" min="1" style="width: 70px;"></label>
                <label>Cols: <input type="number" id="scanColsInput" class="styled-input" value="${state.gridCols}" min="1" style="width: 70px;"></label>
            </div>
            <div style="display: flex; align-items: center; gap: 16px; flex-wrap: wrap; margin-bottom: 8px;">
                <div class="property-field checkbox">
                  <input type="checkbox" id="scanHeaderToggle" ${settings.includeHeader ? 'checked' : ''}>
                  <label for="scanHeaderToggle">Header</label>
                </div>
                <div class="property-field checkbox">
                  <input type="checkbox" id="scanFooterToggle" ${settings.includeFooter ? 'checked' : ''}>
                  <label for="scanFooterToggle">Footer</label>
                </div>
            </div>
            <div style="display: flex; align-items: center; gap: 10px; flex-wrap: wrap; margin-bottom: 8px;">
                <input type="checkbox" id="scanSelectableToggle" ${settings.includeSelectable ? 'checked' : ''}> <label for="scanSelectableToggle" style="margin-right: 16px;">Selectable</label>
                <input type="checkbox" id="scanExpandableToggle" ${settings.includeExpandable ? 'checked' : ''}> <label for="scanExpandableToggle">Expandable</label>
            </div>
        `;
        // Insert scan options before the grid in the grid container
        elements.gridContainer.insertBefore(optionsDiv, elements.gridContainer.firstChild);
        elements.scanOptionsContainer = optionsDiv;

        // Add event listeners for the scan options
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
            // Filter cellProperties to keep only valid cells/headers/footers
            for (const key of Array.from(state.cellProperties.keys())) {
              const match = key.match(/^([0-9]+),([0-9]+)$/);
              if (match) {
                const row = parseInt(match[1], 10);
                const col = parseInt(match[2], 10);
                if (row > state.gridRows || col > state.gridCols) {
                  state.cellProperties.delete(key);
                }
              }
              if (key.startsWith('header-')) {
                const col = parseInt(key.split('-')[1], 10);
                if (col > state.gridCols) {
                  state.cellProperties.delete(key);
                }
              }
            }
            updateCellVisuals();
          }
        };

        document.getElementById('scanRowsInput')?.addEventListener('change', updateGridFromInputs);
        document.getElementById('scanColsInput')?.addEventListener('change', updateGridFromInputs);

        // Add listeners for header/footer toggles
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
      }

      // Select all cells and confirm size
      document.querySelectorAll('.cell').forEach(cell => {
        const key = `${(cell as HTMLDivElement).dataset.row},${(cell as HTMLDivElement).dataset.col}`;
        state.selectedCells.add(key);
        cell.classList.add('selected');
      });
      state.sizeConfirmed = true;

      // Change button text to "Update" and enable it
      elements.createTableBtn.textContent = 'Update Table';
      elements.createTableBtn.disabled = false;
      console.log(`[UI] Set button text to: ${elements.createTableBtn.textContent}, disabled: ${elements.createTableBtn.disabled}`);
      console.log(`[UI] State: hasComponent=${state.hasComponent}, sizeConfirmed=${state.sizeConfirmed}`);

      // Ensure action buttons are visible
      updateModeDependentVisibility();

      setMode('edit');
      updateCreateButtonState();
      break;

    case "table-updated":
      hideLoader();
      showMessage("Table updated successfully!", "success");
      // Keep the button as "Update Table" since we're still editing the same table
      elements.createTableBtn.textContent = 'Update Table';
      break;

    case "component-info":
      console.log(`[UI] Received component info:`, msg.component);
      if (msg.component) {
        state.selectedComponent = msg.component;
        // Now update the visuals with the component information
        updateCellVisuals();
      } else {
        console.log(`[UI] No component info received, using fallback for cell properties`);
        // If no component info, we'll still try to display properties from cellProperties
        updateCellVisuals();
      }
      break;

    case "watsonx-settings":
      {
        const { apiKeyMasked, endpoint, fullApiKey } = msg;
        const endpointInput = document.getElementById('watsonxEndpoint') as HTMLInputElement | null;
        const watsonxApiKeyInput = document.getElementById('watsonxApiKey') as HTMLInputElement | null;
        const spApiKeyInput = document.getElementById('spApiKey') as HTMLInputElement | null;

        if (endpointInput && endpoint) endpointInput.value = endpoint;

        // If we have a full API key, populate both fields
        if (fullApiKey) {
          if (watsonxApiKeyInput) watsonxApiKeyInput.value = fullApiKey;
          if (spApiKeyInput) spApiKeyInput.value = fullApiKey;
        } else if (apiKeyMasked) {
          // Otherwise show masked version as placeholder
          if (watsonxApiKeyInput) watsonxApiKeyInput.placeholder = apiKeyMasked;
          if (spApiKeyInput) spApiKeyInput.placeholder = apiKeyMasked;
        }
      }
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



// Parse CSV file
function parseCSV(file: File): Promise<any[][]> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const text = e.target?.result as string;
        // Split lines and filter out empty lines
        const lines = text.split(/\r?\n/).filter(line => line.trim() !== '');
        const data = lines.map(line => {
          // Simple CSV parsing (doesn't handle quoted fields with commas)
          return line.split(',').map(field => field.trim());
        });
        resolve(data);
      } catch (error) {
        reject(new Error('Failed to parse CSV file'));
      }
    };
    reader.onerror = () => reject(new Error('Failed to read file'));
    reader.readAsText(file);
  });
}

// Parse Excel file
async function parseExcel(file: File): Promise<any[][]> {
  return new Promise((resolve, reject) => {
    // Check if XLSX is already loaded
    // @ts-ignore
    if (typeof window.XLSX !== 'undefined') {
      // @ts-ignore
      processExcelFile(window.XLSX, file, resolve, reject);
      return;
    }

    // Dynamically load SheetJS library
    const script = document.createElement('script');
    script.src = 'https://cdn.sheetjs.com/xlsx-0.20.0/package/dist/xlsx.full.min.js';
    script.onload = () => {
      // Small delay to ensure library is fully loaded
      setTimeout(() => {
        // @ts-ignore
        if (typeof window.XLSX !== 'undefined') {
          // @ts-ignore
          processExcelFile(window.XLSX, file, resolve, reject);
        } else {
          reject(new Error('Failed to load Excel parsing library - library not available after loading'));
        }
      }, 100);
    };
    script.onerror = () => reject(new Error('Failed to load Excel parsing library from CDN'));
    document.head.appendChild(script);
  });
}

// Helper function to process Excel file with XLSX library
function processExcelFile(XLSX: any, file: File, resolve: Function, reject: Function) {
  try {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });

        // Check if workbook has sheets
        if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
          resolve([]); // Return empty array for empty workbook
          return;
        }

        // Get the first worksheet
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        // Convert to JSON
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        // Handle case where sheet_to_json returns empty data
        if (!jsonData || jsonData.length === 0) {
          resolve([]);
          return;
        }

        // Filter out empty rows
        const filteredData = jsonData.filter((row: any[]) =>
          row && row.length > 0 && row.some(cell => cell !== null && cell !== undefined && cell !== '')
        );

        resolve(filteredData);
      } catch (error) {
        reject(new Error('Failed to parse Excel file: ' + (error as Error).message));
      }
    };
    reader.onerror = () => reject(new Error('Failed to read Excel file'));
    reader.readAsArrayBuffer(file);
  } catch (error) {
    reject(new Error('Failed to process Excel file: ' + (error as Error).message));
  }
}

// Parse JSON file
function parseJSON(file: File): Promise<any[][]> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const content = e.target?.result as string;
        const data = JSON.parse(content);

        // Handle different JSON structures
        if (Array.isArray(data)) {
          if (data.length === 0) {
            resolve([]);
            return;
          }

          // If array of objects, convert to array of arrays
          if (typeof data[0] === 'object' && data[0] !== null) {
            // Get all unique keys
            const keys = Array.from(new Set(data.flatMap(Object.keys)));

            // Create header row
            const result = [keys];

            // Add data rows
            data.forEach(obj => {
              const row = keys.map(key => obj[key] ?? '');
              result.push(row);
            });

            resolve(result);
          } else {
            // Already an array of arrays
            resolve(data);
          }
        } else if (typeof data === 'object' && data !== null) {
          // Handle object with columns and rows structure
          if (Array.isArray(data.columns) && Array.isArray(data.rows)) {
            // Create header row from columns
            const result = [data.columns];

            // Add data rows
            data.rows.forEach((obj: any) => {
              const row = data.columns.map((col: string) => obj[col] ?? '');
              result.push(row);
            });

            resolve(result);
          } else {
            // Handle generic object structure
            const keys = Object.keys(data);
            if (keys.length > 0) {
              // Try to convert to array format
              const firstKey = keys[0];
              const firstValue = data[firstKey];

              if (Array.isArray(firstValue)) {
                // Assume all values are arrays of the same length
                const result = [keys];
                for (let i = 0; i < firstValue.length; i++) {
                  const row = keys.map(key => data[key][i] ?? '');
                  result.push(row);
                }
                resolve(result);
              } else {
                reject(new Error('Unsupported JSON structure'));
              }
            } else {
              reject(new Error('JSON object is empty'));
            }
          }
        } else {
          reject(new Error('JSON file must contain an array or object'));
        }
      } catch (error) {
        console.error('JSON parsing error:', error);
        reject(new Error('Failed to parse JSON file'));
      }
    };
    reader.onerror = () => reject(new Error('Failed to read file'));
    reader.readAsText(file);
  });
}

// Show file preview
function showFilePreview(data: any[][]) {
  // Update the main grid with file data
  updateMainGridWithData(data);

  // Show the inline remove button
  const removeFileBtn = document.getElementById('removeFileBtn');
  if (removeFileBtn) {
    removeFileBtn.style.display = 'flex';
  }
}

// Reset grid to default state
function resetGridToDefault() {
  // Clear all cell properties
  state.cellProperties.clear();

  // Recreate the grid with default values
  createGrid();

  // Update cell visuals
  updateCellVisuals();

  // Show success message
  showMessage('File data removed and grid reset to default state', 'success');
}

// Update main grid with file data
function updateMainGridWithData(data: any[][]) {
  // Hide file preview container
  const previewContainer = document.getElementById('filePreview');
  if (previewContainer) {
    previewContainer.style.display = 'none';
  }

  // Get the number of rows and columns in the data
  const dataRows = data.length;
  const dataCols = data.length > 0 ? data[0].length : 0;

  if (dataRows === 0 || dataCols === 0) {
    showMessage('No data found in file', 'error');
    return;
  }

  // Update grid dimensions if needed
  const rowsInput = document.getElementById('scanRowsInput') as HTMLInputElement | null;
  const colsInput = document.getElementById('scanColsInput') as HTMLInputElement | null;

  if (rowsInput && colsInput) {
    // Set new dimensions (account for header row)
    const newRows = Math.max(1, dataRows - 1); // Subtract 1 for header row
    const newCols = Math.max(1, dataCols);

    rowsInput.value = String(newRows);
    colsInput.value = String(newCols);

    // Update state and grid
    state.gridRows = newRows;
    state.gridCols = newCols;
    createGrid();
  }

  // Process header row if it exists
  if (data.length > 0) {
    const headerRow = data[0];
    for (let c = 0; c < headerRow.length && c < state.gridCols; c++) {
      const headerKey = `header-${c + 1}`;
      const text = String(headerRow[c] || '');

      // Update state with header text
      if (!state.cellProperties.has(headerKey)) {
        state.cellProperties.set(headerKey, {
          properties: {},
          type: 'HEADER'
        });
      }

      const headerProps = state.cellProperties.get(headerKey);
      if (headerProps) {
        // Use the same approach as AI data - find the actual text property from the component
        if (state.headerCellComponent?.availableProperties && state.headerCellComponent?.propertyTypes) {
          const textProp = state.headerCellComponent.availableProperties.find(p =>
            state.headerCellComponent!.propertyTypes[p] === 'TEXT');
          if (textProp) {
            headerProps.properties[textProp] = text;
          } else {
            // Fallback to hardcoded property name if we can't find it
            headerProps.properties['Cell text#12234:32'] = text;
          }
        } else {
          // Fallback to hardcoded property name if we don't have component info
          headerProps.properties['Cell text#12234:32'] = text;
        }
      }
    }
  }

  // Process data rows
  for (let r = 1; r < data.length && r <= state.gridRows; r++) {
    const row = data[r];
    for (let c = 0; c < row.length && c < state.gridCols; c++) {
      // Use the correct key format that matches what createTable expects
      const cellKey = `${r},${c + 1}`;  // UI format (1-based)
      const backendKey = `${r - 1}-${c}`;  // Backend format (0-based)
      const text = String(row[c] || '');

      // Update state with cell text using the UI key format
      if (!state.cellProperties.has(cellKey)) {
        state.cellProperties.set(cellKey, {
          properties: {},
          type: 'CELL'
        });
      }

      const cellProps = state.cellProperties.get(cellKey);
      if (cellProps) {
        // Use the same approach as AI data - find the actual text property from the component
        if (state.selectedComponent?.availableProperties && state.selectedComponent?.propertyTypes) {
          const textProp = state.selectedComponent.availableProperties.find(p =>
            state.selectedComponent!.propertyTypes[p] === 'TEXT');
          if (textProp) {
            cellProps.properties[textProp] = text;
          } else {
            // Fallback to hardcoded property name if we can't find it
            cellProps.properties['Cell text#12234:32'] = text;
          }
        } else {
          // Fallback to hardcoded property name if we don't have component info
          cellProps.properties['Cell text#12234:32'] = text;
        }
      }
    }
  }

  // Recreate the grid to properly display all values
  createGrid();

  // Ensure header cells are updated with the new data
  setTimeout(() => {
    renderHeaderFooterGrids();
  }, 0);

  markChangesForReset(); // Enable reset button after file data is loaded
  showMessage('File data loaded successfully', 'success');
}

function sortColumnData(cellProperties: Map<string, any>, cols: number, rows: number): Map<string, any> {
  let hasSortableColumns = false;
  
  // Quick check if any columns need sorting
  for (let c = 1; c <= cols; c++) {
    const headerData = cellProperties.get(`header-${c}`);
    if (headerData?.properties?.['Sortable'] === 'True' && 
        (headerData.properties['Sorted'] === 'Ascending' || headerData.properties['Sorted'] === 'Descending')) {
      hasSortableColumns = true;
      break;
    }
  }
  
  // Return original if no sorting needed
  if (!hasSortableColumns) return cellProperties;
  
  const sortedCellProperties = new Map(cellProperties);
  
  for (let c = 1; c <= cols; c++) {
    const headerData = sortedCellProperties.get(`header-${c}`);
    const sortable = headerData?.properties?.['Sortable'];
    const sorted = headerData?.properties?.['Sorted'];
    
    if (sortable === 'True' && (sorted === 'Ascending' || sorted === 'Descending')) {
      const columnData: { rowIndex: number; cellData: any; value: string }[] = [];
      
      for (let r = 1; r <= rows; r++) {
        const cellData = sortedCellProperties.get(`${r},${c}`);
        let cellValue = '';
        
        if (cellData?.properties) {
          const textProp = Object.keys(cellData.properties).find(prop => 
            prop.toLowerCase().includes('text') && !prop.toLowerCase().includes('second') &&
            typeof cellData.properties[prop] === 'string');
          cellValue = (textProp && cellData.properties[textProp]) || 
                     cellData.properties['Cell text#12234:32'] || '';
        }
        
        columnData.push({ rowIndex: r, cellData, value: String(cellValue) });
      }
      
      columnData.sort((a, b) => {
        const numA = parseFloat(a.value);
        const numB = parseFloat(b.value);
        
        if (!isNaN(numA) && !isNaN(numB)) {
          return sorted === 'Ascending' ? numA - numB : numB - numA;
        }
        return sorted === 'Ascending' ? a.value.localeCompare(b.value) : b.value.localeCompare(a.value);
      });
      
      for (let r = 1; r <= rows; r++) {
        sortedCellProperties.set(`${r},${c}`, columnData[r - 1].cellData);
      }
    }
  }
  
  return sortedCellProperties;
}

// Column Reordering Functions
function openColumnReorderModal() {
  const overlay = document.getElementById('columnReorderOverlay');
  const columnList = document.getElementById('columnList');
  
  if (!overlay || !columnList) return;
  
  // Clear existing column items
  columnList.innerHTML = '';
  
  // Create column items based on current grid columns
  for (let c = 1; c <= state.gridCols; c++) {
    const columnItem = createColumnItem(c);
    columnList.appendChild(columnItem);
  }
  
  // Show modal
  overlay.style.display = 'flex';
  
  // Setup modal event listeners
  setupColumnReorderModalListeners();
}

function createColumnItem(columnIndex: number): HTMLElement {
  const item = document.createElement('div');
  item.className = 'column-item';
  item.draggable = true;
  item.dataset.columnIndex = String(columnIndex);
  
  // Get column name from header if available
  const headerKey = `header-${columnIndex}`;
  const headerData = state.cellProperties.get(headerKey);
  let columnName = `Column ${columnIndex}`;
  
  if (headerData && headerData.properties) {
    // Try to find text property
    const textProp = Object.keys(headerData.properties).find(prop => 
      prop.toLowerCase().includes('text') && 
      typeof headerData.properties[prop] === 'string' &&
      headerData.properties[prop].trim() !== ''
    );
    
    if (textProp) {
      columnName = headerData.properties[textProp];
    }
  }
  
  item.innerHTML = `
    <span class="drag-handle">⋮⋮</span>
    <span class="column-name">${columnName}</span>
    <span class="column-index">${columnIndex}</span>
    <button class="column-delete-btn" title="Delete column">×</button>
  `;
  
  // Add delete functionality
  const deleteBtn = item.querySelector('.column-delete-btn') as HTMLButtonElement;
  deleteBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    item.classList.toggle('deleted');
    item.draggable = !item.classList.contains('deleted');
  });
  
  // Add drag event listeners
  item.addEventListener('dragstart', handleDragStart);
  item.addEventListener('dragover', handleDragOver);
  item.addEventListener('drop', handleDrop);
  item.addEventListener('dragend', handleDragEnd);
  
  return item;
}

function setupColumnReorderModalListeners() {
  const closeBtn = document.getElementById('closeReorderModal');
  const cancelBtn = document.getElementById('cancelReorderBtn');
  const applyBtn = document.getElementById('applyReorderBtn');
  const overlay = document.getElementById('columnReorderOverlay');
  
  // Close modal handlers
  const closeModal = () => {
    if (overlay) overlay.style.display = 'none';
  };
  
  closeBtn?.addEventListener('click', closeModal);
  cancelBtn?.addEventListener('click', closeModal);
  
  // Apply reorder handler
  applyBtn?.addEventListener('click', applyColumnReorder);
  
  // Close on overlay click
  overlay?.addEventListener('click', (e) => {
    if (e.target === overlay) closeModal();
  });
}

let draggedElement: HTMLElement | null = null;

function handleDragStart(e: DragEvent) {
  draggedElement = e.target as HTMLElement;
  draggedElement.classList.add('dragging');
  
  if (e.dataTransfer) {
    e.dataTransfer.effectAllowed = 'move';
    e.dataTransfer.setData('text/html', draggedElement.outerHTML);
  }
}

function handleDragOver(e: DragEvent) {
  e.preventDefault();
  
  const target = e.target as HTMLElement;
  const columnItem = target.closest('.column-item') as HTMLElement;
  
  if (columnItem && columnItem !== draggedElement) {
    columnItem.classList.add('drag-over');
  }
}

function handleDrop(e: DragEvent) {
  e.preventDefault();
  
  const target = e.target as HTMLElement;
  const columnItem = target.closest('.column-item') as HTMLElement;
  
  if (columnItem && draggedElement && columnItem !== draggedElement) {
    const columnList = document.getElementById('columnList');
    if (columnList) {
      // Get all column items
      const items = Array.from(columnList.children) as HTMLElement[];
      const draggedIndex = items.indexOf(draggedElement);
      const targetIndex = items.indexOf(columnItem);
      
      // Reorder the items
      if (draggedIndex < targetIndex) {
        columnList.insertBefore(draggedElement, columnItem.nextSibling);
      } else {
        columnList.insertBefore(draggedElement, columnItem);
      }
    }
  }
  
  // Clean up drag over styles
  document.querySelectorAll('.column-item').forEach(item => {
    item.classList.remove('drag-over');
  });
}

function handleDragEnd(e: DragEvent) {
  const target = e.target as HTMLElement;
  target.classList.remove('dragging');
  
  // Clean up all drag styles
  document.querySelectorAll('.column-item').forEach(item => {
    item.classList.remove('drag-over', 'dragging');
  });
  
  draggedElement = null;
}

function applyColumnReorder() {
  const columnList = document.getElementById('columnList');
  if (!columnList) return;
  
  // Get remaining columns (not deleted) and their new order
  const columnItems = Array.from(columnList.children) as HTMLElement[];
  const remainingColumns = columnItems
    .filter(item => !item.classList.contains('deleted'))
    .map(item => parseInt(item.dataset.columnIndex || '0'));
  
  console.log('Remaining columns:', remainingColumns);
  
  // Apply the reordering and deletion to the table data
  reorderAndDeleteColumns(remainingColumns);
  
  // Close modal
  const overlay = document.getElementById('columnReorderOverlay');
  if (overlay) overlay.style.display = 'none';
  
  // Show success message
  const deletedCount = columnItems.length - remainingColumns.length;
  const message = deletedCount > 0 
    ? `Columns reordered and ${deletedCount} column(s) deleted successfully!`
    : 'Columns reordered successfully!';
  showMessage(message, 'success');
}

function reorderAndDeleteColumns(remainingColumns: number[]) {
  // Update grid column count
  const newColCount = remainingColumns.length;
  state.gridCols = newColCount;
  
  // Create new cell properties map with reordered and filtered columns
  const newCellProperties = new Map<string, any>();
  
  // Reorder and filter header cells
  remainingColumns.forEach((oldCol, newIndex) => {
    const newCol = newIndex + 1; // 1-based indexing
    const oldKey = `header-${oldCol}`;
    const newKey = `header-${newCol}`;
    const headerData = state.cellProperties.get(oldKey);
    if (headerData) {
      newCellProperties.set(newKey, headerData);
    }
  });
  
  // Reorder and filter body cells
  for (let row = 1; row <= state.gridRows; row++) {
    remainingColumns.forEach((oldCol, newIndex) => {
      const newCol = newIndex + 1; // 1-based indexing
      const oldKey = `${row},${oldCol}`;
      const newKey = `${row},${newCol}`;
      const cellData = state.cellProperties.get(oldKey);
      if (cellData) {
        newCellProperties.set(newKey, cellData);
      }
    });
  }
  
  // Keep footer data as is
  const footerData = state.cellProperties.get('footer');
  if (footerData) {
    newCellProperties.set('footer', footerData);
  }
  
  // Update state with reordered and filtered data
  state.cellProperties = newCellProperties;
  
  // Update scan options if they exist
  const colsInput = document.getElementById('scanColsInput') as HTMLInputElement | null;
  if (colsInput) {
    colsInput.value = String(newColCount);
  }
  
  // Recreate the grid to reflect the new structure
  createGrid();
  
  // Mark changes for reset
  markChangesForReset();
}

// Make openColumnReorderModal available globally
(window as any).openColumnReorderModal = openColumnReorderModal;

export { }; // Treat this file as a module 