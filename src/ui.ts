import { fakerMethods } from './faker-methods';
import { init as initAnalytics, track as trackEvent } from './utils/analytics';

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
  // collectCarbonKeysBtn: HTMLButtonElement;
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
  // slotCheckbox: HTMLInputElement; // Removed - using dynamic slot checkbox only
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
  // elements.collectCarbonKeysBtn = document.getElementById('collectCarbonKeysBtn') as HTMLButtonElement;
  elements.propertyEditor = document.getElementById('propertyEditor')!;
  
  // Request current selection check on plugin load
  // This handles the case where a table is already selected when the plugin opens
  console.log('[UI] Plugin UI loaded, requesting selection check...');
  parent.postMessage({ pluginMessage: { type: 'check-current-selection' } }, '*');
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
  // elements.slotCheckbox = document.getElementById('slotCheckbox') as HTMLInputElement; // Removed - using dynamic slot checkbox only
  elements.customCellTextToggle = document.getElementById('customCellTextToggle') as HTMLInputElement;
  elements.customCellText = document.getElementById('customCellText') as HTMLInputElement;
  elements.customCellTextContainer = document.getElementById('customCellTextContainer')!;
  elements.landingPage = document.getElementById('landingPage')!;
  elements.componentModeBtn = document.getElementById('componentModeBtn') as HTMLButtonElement | null;
  elements.propertyEditorOverlay = document.getElementById('propertyEditorOverlay')!;
  elements.componentDisplay = document.getElementById('componentDisplay')!;
  elements.scanOptionsContainer = null;

  domReady = true;

  // Initialize reset button as disabled
  updateResetButtonState();

  // REMOVED: Automatic scan-table on plugin load
  // This was causing issues when opening the plugin with a generated table selected
  // The check-current-selection message (sent above) will handle the proper flow
  // parent.postMessage({ pluginMessage: { type: 'scan-table' } }, '*');

  // Now safe to call setup functions
  createGrid();

    // Wire main-page single prompt controls
    const useSinglePrompt = document.getElementById('useSinglePrompt') as HTMLInputElement | null;
    const singlePromptContainer = document.getElementById('singlePromptContainer') as HTMLElement | null;
    const generateTableFromPromptBtn = document.getElementById('generateTableFromPromptBtn') as HTMLButtonElement | null;
    useSinglePrompt?.addEventListener('change', () => {
        if (singlePromptContainer) {
            if (useSinglePrompt.checked) {
                singlePromptContainer.style.display = 'block';
        requestAnimationFrame(() => {
                singlePromptContainer.classList.add('show');
        });
            } else {
                singlePromptContainer.classList.remove('show');
        setTimeout(() => {
          singlePromptContainer.style.display = 'none';
        }, 250);
            }
        }

        // If single prompt is selected, deselect file upload
        const useFileUpload = document.getElementById('useFileUpload') as HTMLInputElement | null;
        const fileUploadContainer = document.getElementById('fileUploadContainer') as HTMLElement | null;
        if (useSinglePrompt?.checked && useFileUpload) {
            useFileUpload.checked = false;
            if (fileUploadContainer) {
                fileUploadContainer.classList.remove('show');
        setTimeout(() => {
          fileUploadContainer.style.display = 'none';
        }, 250);
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
      generateTableFromPromptBtn.title = hasPrompt ? 'Generate table with AI' : 'Enter a prompt to enable';
      
      // Add visual feedback for character count
      const charCount = singlePromptText.value.length;
      if (charCount > 0 && charCount < 10) {
        generateTableFromPromptBtn.title = 'Prompt too short - add more details';
      }
    }
  }

  // Initially disable the button with better visual feedback
  if (generateTableFromPromptBtn) {
    generateTableFromPromptBtn.disabled = true;
    generateTableFromPromptBtn.style.opacity = '0.5';
    generateTableFromPromptBtn.style.cursor = 'not-allowed';
    generateTableFromPromptBtn.title = 'Enter a prompt to enable';
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
            if (useFileUpload.checked) {
                fileUploadContainer.style.display = 'block';
        requestAnimationFrame(() => {
                fileUploadContainer.classList.add('show');
        });
            } else {
                fileUploadContainer.classList.remove('show');
        setTimeout(() => {
          fileUploadContainer.style.display = 'none';
        }, 250);
            }
        }

        // If file upload is selected, deselect single prompt
        const useSinglePrompt = document.getElementById('useSinglePrompt') as HTMLInputElement | null;
        const singlePromptContainer = document.getElementById('singlePromptContainer') as HTMLElement | null;
        if (useFileUpload?.checked && useSinglePrompt) {
            useSinglePrompt.checked = false;
            if (singlePromptContainer) {
                singlePromptContainer.classList.remove('show');
        setTimeout(() => {
          singlePromptContainer.style.display = 'none';
        }, 250);
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

    // Add loading state to button
    generateTableFromPromptBtn.classList.add('loading');
    generateTableFromPromptBtn.disabled = true;
    
    // Show loader when starting AI generation
    showLoader('Generating content with AI...');

    // Only using watsonx.ai for single prompt generation
    // API key validation removed - now using Code Engine proxy server
    const apiKey = ''; // Not needed when using proxy server

    // Debug: Log what we're sending to the backend (sensitive data removed)
    console.log('=== DEBUG: UI sending generate-table-with-ai message ===');
    console.log('Rows:', rows);
    console.log('Cols:', cols);
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

  // Setup landing page button logic (only if element exists)
  if (elements.componentModeBtn) {
    elements.componentModeBtn.onclick = () => {
      elements.landingPage.style.display = 'none';
    };
  }

  // Initialize tooltip
  state.tooltip.className = 'cell-tooltip';

  // Initialize analytics once DOM is ready (after elements are wired)
  try {
		initAnalytics({ endpoint: 'https://application-e9.21hwt6k1vujm.us-east.codeengine.appdomain.cloud/analytics', enabled: true });
		trackEvent('plugin_open', { version: '1.0.0' });
	} catch (e) {
		console.warn('[analytics] init failure', e);
	}

  // Track single-prompt generate button click with prompt text
  if (generateTableFromPromptBtn) {
		generateTableFromPromptBtn.addEventListener('click', () => {
			const promptVal = (document.getElementById('singlePromptText') as HTMLTextAreaElement | null)?.value?.trim() || '';
			trackEvent('ai_generate_single', { prompt: promptVal, length: promptVal.length });
		});
	}
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

  console.log(`[DEBUG] createGrid - Before sorting: gridCols=${state.gridCols}, gridRows=${state.gridRows}`);
  console.log(`[DEBUG] createGrid - cellProperties keys:`, Array.from(state.cellProperties.keys()));
  
  const sortedCellProperties = sortColumnData(state.cellProperties, state.gridCols, state.gridRows);
  // Update state with sorted data to ensure consistency
  state.cellProperties = sortedCellProperties;
  
  console.log(`[DEBUG] createGrid - After sorting: cellProperties keys:`, Array.from(state.cellProperties.keys()));

  let cellIndex = 0;
  for (let r = 1; r <= state.gridRows; r++) {
    for (let c = 1; c <= state.gridCols; c++) {
      const cell = document.createElement('div');
      cell.className = 'cell';
      cell.dataset.row = String(r);
      cell.dataset.col = String(c);
      
      // Add staggered animation delay
      cell.style.animationDelay = `${cellIndex * 20}ms`;
      cellIndex++;

      const key = `${r},${c}`;
      const cellState = sortedCellProperties.get(key);
      console.log(`[DEBUG] createGrid - Cell ${key}: cellState=`, cellState);
      let displayText = '';
      let hasSlotEnabled = false;
      
      if (cellState?.properties && state.selectedComponent) {
        // Check for slot using the actual component property as source of truth
        const slotPropName = state.selectedComponent.availableProperties?.find(p => 
          state.selectedComponent!.propertyTypes[p] === 'BOOLEAN' && 
          p.toLowerCase().includes('slot') && 
          !p.toLowerCase().includes('swap')
        );
        
        if (slotPropName && cellState.properties[slotPropName] !== undefined) {
          hasSlotEnabled = cellState.properties[slotPropName] === true;
        } else {
          // Fallback to dedicated slot variable if property doesn't exist
          hasSlotEnabled = cellState.slot || false;
        }
        
        if (hasSlotEnabled) {
          cell.classList.add('slot-enabled');
        }

        // Only use slot component data if slot is actually enabled
        if (hasSlotEnabled && cellState.slotComponentProps) {
          if (cellState.slotComponentProps.userName) {
            displayText = cellState.slotComponentProps.userName;
          } else if (cellState.slotComponentProps.tagText) {
            displayText = cellState.slotComponentProps.tagText;
          } else if (cellState.slotComponentProps.overflowActions) {
            displayText = cellState.slotComponentProps.overflowActions;
          } else if (cellState.slotComponentProps.editAction) {
            displayText = cellState.slotComponentProps.editAction;
          } else if (cellState.slotComponentProps.deleteAction) {
            displayText = cellState.slotComponentProps.deleteAction;
          }
        }
        
        // If slot is disabled or no slot component data, use regular text properties
        if (!displayText) {
          const textProp = state.selectedComponent.availableProperties?.find(p =>
            state.selectedComponent!.propertyTypes[p] === 'TEXT');
          displayText = (textProp && cellState.properties[textProp]) || 
                       cellState.properties['Cell text#12234:32'] || '';
        }
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

  // Update selectedCells based on new apply mode
  if (state.currentEditingCell && !state.currentEditingCell.startsWith('header-') && state.currentEditingCell !== 'footer') {
    const [row, col] = state.currentEditingCell.split(',').map(Number);
    state.selectedCells.clear();
    
    if (newApplyMode === 'cell') {
      state.selectedCells.add(state.currentEditingCell);
      console.log(`[handleApplyOptionClick] Updated selectedCells: single cell ${state.currentEditingCell}`);
    } else if (newApplyMode === 'row') {
      for (let c = 1; c <= state.gridCols; c++) {
        state.selectedCells.add(`${row},${c}`);
      }
      console.log(`[handleApplyOptionClick] Updated selectedCells: ${state.selectedCells.size} cells in row ${row}`);
    } else if (newApplyMode === 'column') {
      for (let r = 1; r <= state.gridRows; r++) {
        state.selectedCells.add(`${r},${col}`);
      }
      console.log(`[handleApplyOptionClick] Updated selectedCells: ${state.selectedCells.size} cells in column ${col}`);
    }
  }

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
  // elements.collectCarbonKeysBtn.addEventListener('click', collectCarbonKeys);
  elements.cancelPropsBtn.addEventListener('click', closePropertyEditor);
  elements.propertyEditorOverlay.addEventListener('click', (e) => {
    if (e.target === elements.propertyEditorOverlay) {
      closePropertyEditor();
    }
  });
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
  
  // Removed duplicate slot checkbox event listener - now handled in renderSlotComponentControls
  document.querySelectorAll('.apply-option').forEach(option => {
    option.addEventListener('click', handleApplyOptionClick);
  });
  // Helper to toggle which AI config is visible
  function updateAiConfigVisibility() {
    const aiSource = document.getElementById('aiSource') as HTMLSelectElement | null;
    const fakerConfig = document.getElementById('fakerConfig') as HTMLElement | null;
    const watsonxConfig = document.getElementById('watsonxConfig') as HTMLElement | null;
    const isChecked = elements.generateSampleCheckbox.checked;

    if (!aiSource || (!fakerConfig && !watsonxConfig)) return;

    if (!isChecked) {
      // Hide both when AI generation is disabled
      if (fakerConfig) { fakerConfig.classList.remove('show'); fakerConfig.style.display = 'none'; }
      if (watsonxConfig) { watsonxConfig.classList.remove('show'); watsonxConfig.style.display = 'none'; }
      return;
    }

    // Show only the selected source config
    const v = aiSource.value;
    if (fakerConfig) {
      if (v === 'faker') {
        fakerConfig.style.display = 'block';
        requestAnimationFrame(() => { fakerConfig.classList.add('show'); });
      } else {
        fakerConfig.classList.remove('show');
        setTimeout(() => { fakerConfig.style.display = 'none'; }, 250);
      }
    }
    if (watsonxConfig) {
      if (v === 'watsonx') {
        watsonxConfig.style.display = 'block';
        requestAnimationFrame(() => { watsonxConfig.classList.add('show'); });
      } else {
        watsonxConfig.classList.remove('show');
        setTimeout(() => { watsonxConfig.style.display = 'none'; }, 250);
      }
    }
  }

  elements.generateSampleCheckbox.addEventListener('change', () => {
    const isChecked = elements.generateSampleCheckbox.checked;
    elements.customCellText.disabled = isChecked;
    
    if (isChecked) {
      elements.aiPromptContainer.style.display = 'block';
      requestAnimationFrame(() => {
        elements.aiPromptContainer.classList.add('show');
      });
      // Ensure the correct AI config is shown immediately on first enable
      updateAiConfigVisibility();
    } else {
      elements.aiPromptContainer.classList.remove('show');
      setTimeout(() => {
        elements.aiPromptContainer.style.display = 'none';
      }, 250);
      // Hide both configs when disabling
      updateAiConfigVisibility();
    }
    
    // Disable custom cell text toggle when AI generation is enabled
    elements.customCellTextToggle.disabled = isChecked;
  });
  // AI source toggle
  // AI source toggle back
  const aiSource = document.getElementById('aiSource') as HTMLSelectElement | null;
  const fakerConfig = document.getElementById('fakerConfig') as HTMLElement | null;
  const watsonxConfig = document.getElementById('watsonxConfig') as HTMLElement | null;
  aiSource?.addEventListener('change', () => {
    updateAiConfigVisibility();
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
    resetBtn.title = hasChangesToReset ? 'Reset all table properties' : 'No changes to reset';
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





// Store original form values when opening property editor
let originalFormValues: any = null;

function closePropertyEditor() {
  // If we have original values stored, restore them before closing
  if (originalFormValues && state.currentEditingCell) {
    restoreFormValues(originalFormValues);
    originalFormValues = null; // Clear the stored values
    
    // Update the preview grid to reflect the restored values
    updateCellVisuals();
  }
  
  elements.propertyEditor.style.opacity = '0';
  elements.propertyEditorOverlay.style.display = 'none';
  setTimeout(() => {
    elements.propertyEditor.style.display = 'none';
    state.currentEditingCell = null;
  }, 300);
}

function captureOriginalFormValues() {
  const values: any = {};
  
  // Capture custom cell text toggle
  if (elements.customCellTextToggle) {
    values.customCellTextEnabled = elements.customCellTextToggle.checked;
  }
  
  // Capture custom cell text value
  if (elements.customCellText) {
    values.customCellText = elements.customCellText.value;
  }
  
  // Capture second text line toggle and value
  if (elements.secondTextLine) {
    values.secondTextLine = elements.secondTextLine.checked;
  }
  if (elements.secondCellText) {
    values.secondCellText = elements.secondCellText.value;
  }
  
  // Capture slot checkbox - now handled dynamically
  const dynamicSlotCheckbox = document.getElementById('dynamicSlotCheckbox') as HTMLInputElement;
  if (dynamicSlotCheckbox) {
    values.slotEnabled = dynamicSlotCheckbox.checked;
  }
  
  // Capture slot component specific fields
  const slotProps: any = {};
  const userNameField = document.getElementById('userName') as HTMLInputElement;
  if (userNameField) slotProps.userName = userNameField.value;
  
  const tagTextField = document.getElementById('tagText') as HTMLInputElement;
  if (tagTextField) slotProps.tagText = tagTextField.value;
  
  const actionsField = document.getElementById('overflowActions') as HTMLInputElement;
  if (actionsField) slotProps.overflowActions = actionsField.value;
  
  const editField = document.getElementById('editAction') as HTMLInputElement;
  if (editField) slotProps.editAction = editField.value;
  
  const deleteField = document.getElementById('deleteAction') as HTMLInputElement;
  if (deleteField) slotProps.deleteAction = deleteField.value;
  
  if (Object.keys(slotProps).length > 0) {
    values.slotComponentProps = slotProps;
  }
  
  // Capture dynamic property fields
  const dynamicProps: any = {};
  document.querySelectorAll('[id^="dynamic-"]').forEach(input => {
    const element = input as HTMLInputElement | HTMLSelectElement;
    const propName = element.id.replace('dynamic-', '');
    if (element.type === 'checkbox') {
      dynamicProps[propName] = (element as HTMLInputElement).checked;
    } else {
      dynamicProps[propName] = element.value;
    }
  });
  
  if (Object.keys(dynamicProps).length > 0) {
    values.dynamicProps = dynamicProps;
  }
  
  console.log('[DEBUG] Captured original form values:', values);
  return values;
}

function restoreFormValues(originalValues: any) {
  // Restore custom cell text toggle
  if (elements.customCellTextToggle && originalValues.customCellTextEnabled !== undefined) {
    elements.customCellTextToggle.checked = originalValues.customCellTextEnabled;
  }
  
  // Restore custom cell text value
  if (elements.customCellText && originalValues.customCellText !== undefined) {
    elements.customCellText.value = originalValues.customCellText;
  }
  
  // Restore second text line toggle and value
  if (elements.secondTextLine && originalValues.secondTextLine !== undefined) {
    elements.secondTextLine.checked = originalValues.secondTextLine;
  }
  if (elements.secondCellText && originalValues.secondCellText !== undefined) {
    elements.secondCellText.value = originalValues.secondCellText;
  }
  
  // Restore slot checkbox and slot component properties - now handled dynamically
  const dynamicSlotCheckbox = document.getElementById('dynamicSlotCheckbox') as HTMLInputElement;
  if (dynamicSlotCheckbox && originalValues.slotEnabled !== undefined) {
    dynamicSlotCheckbox.checked = originalValues.slotEnabled;
  }
  
  // Restore slot component specific fields
  if (originalValues.slotComponentProps) {
    const slotProps = originalValues.slotComponentProps;
    if (slotProps.userName) {
      const userNameField = document.getElementById('userName') as HTMLInputElement;
      if (userNameField) userNameField.value = slotProps.userName;
    }
    if (slotProps.tagText) {
      const tagTextField = document.getElementById('tagText') as HTMLInputElement;
      if (tagTextField) tagTextField.value = slotProps.tagText;
    }
    if (slotProps.overflowActions) {
      const actionsField = document.getElementById('overflowActions') as HTMLInputElement;
      if (actionsField) actionsField.value = slotProps.overflowActions;
    }
    if (slotProps.editAction) {
      const editField = document.getElementById('editAction') as HTMLInputElement;
      if (editField) editField.value = slotProps.editAction;
    }
    if (slotProps.deleteAction) {
      const deleteField = document.getElementById('deleteAction') as HTMLInputElement;
      if (deleteField) deleteField.value = slotProps.deleteAction;
    }
  }
  
  // Restore dynamic property fields
  if (originalValues.dynamicProps) {
    Object.entries(originalValues.dynamicProps).forEach(([propName, value]) => {
      const input = document.getElementById(`dynamic-${propName}`) as HTMLInputElement | HTMLSelectElement;
      if (input) {
        if (input.type === 'checkbox') {
          (input as HTMLInputElement).checked = !!value;
        } else {
          input.value = String(value || '');
        }
      }
    });
  }
  
  console.log('[DEBUG] Form values restored from original state');
}

function getCellState(key: string) {
  console.log(`[DEBUG] getCellState called for key: ${key}`);
  console.log(`[DEBUG] getCellState - state.cellProperties exists:`, !!state.cellProperties);
  console.log(`[DEBUG] getCellState - state.cellProperties type:`, typeof state.cellProperties);
  
  if (!state.cellProperties) {
    console.error(`[DEBUG] getCellState - state.cellProperties is undefined!`);
    return { properties: {}, slot: false };
  }
  
  if (!state.cellProperties.has(key)) {
    console.log(`[DEBUG] getCellState - creating new cell state for key: ${key}`);
    state.cellProperties.set(key, { 
      properties: {},
      slot: false  // Initialize slot state as false
    });
  }
  
  const result = state.cellProperties.get(key);
  console.log(`[DEBUG] getCellState - returning:`, result);
  
  // If result is undefined, create a new cell state
  if (!result) {
    console.log(`[DEBUG] getCellState - result was undefined, creating new cell state for key: ${key}`);
    const newCellState = { 
      properties: {},
      slot: false
    };
    state.cellProperties.set(key, newCellState);
    return newCellState;
  }
  
  return result;
}

// Function to determine default column width based on content type (UI version)
function getDefaultColumnWidthUI(columnIndex: number): number {
  // Check header text to determine column type
  const headerKey = `header-${columnIndex}`;
  const headerData = state.cellProperties.get(headerKey);
  let headerText = '';
  
  if (headerData && headerData.properties) {
    // Try to find text property in header
    const textProps = Object.keys(headerData.properties).filter(prop => 
      prop.toLowerCase().includes('text') && 
      typeof headerData.properties[prop] === 'string' &&
      headerData.properties[prop].trim() !== ''
    );
    if (textProps.length > 0) {
      headerText = headerData.properties[textProps[0]].toLowerCase();
    }
  }
  
  // Check column content for action indicators
  let hasActionContent = false;
  let hasFullNameContent = false;
  
  // Sample a few cells to determine content type
  for (let r = 1; r <= Math.min(3, state.gridRows); r++) { // Check first 3 rows
    const cellKey = `${r},${columnIndex}`;
    const cellData = state.cellProperties.get(cellKey);
    if (cellData && cellData.properties) {
      const textProps = Object.keys(cellData.properties).filter(prop => 
        prop.toLowerCase().includes('text') && 
        typeof cellData.properties[prop] === 'string'
      );
      
      for (const textProp of textProps) {
        const cellText = cellData.properties[textProp].toLowerCase();
        
        // Check for action indicators
        if (cellText.includes('edit') || cellText.includes('delete') || 
            cellText.includes('view') || cellText.includes('action') ||
            cellText.includes('manage') || cellText.includes('options')) {
          hasActionContent = true;
        }
        
        // Check for full name indicators
        if (cellText.includes(' ') && cellText.length > 10 && 
            /^[a-zA-Z\s]+$/.test(cellText)) {
          hasFullNameContent = true;
        }
      }
    }
  }
  
  // Determine width based on content type
  if (hasActionContent || headerText.includes('action') || headerText.includes('manage')) {
    console.log(`[getDefaultColumnWidthUI] Column ${columnIndex} detected as ACTION column - using 80px width`);
    return 80;
  } else if (hasFullNameContent || headerText.includes('name') || headerText.includes('full')) {
    console.log(`[getDefaultColumnWidthUI] Column ${columnIndex} detected as FULL NAME column - using 160px width`);
    return 160;
  } else {
    console.log(`[getDefaultColumnWidthUI] Column ${columnIndex} using default 120px width`);
    return 120;
  }
}

function initializeSlotVariables() {
  console.log('[initializeSlotVariables] Initializing slot variables for all cells');
  
  // Find the swap slot property from available properties
  let swapSlotProp = state.selectedComponent?.availableProperties?.find((p: string) => 
    p.toLowerCase().includes('swap') && p.toLowerCase().includes('slot')
  );
  
  // Initialize slot variables for all cells
  for (let r = 1; r <= state.gridRows; r++) {
    for (let c = 1; c <= state.gridCols; c++) {
      const key = `${r},${c}`;
      const cellState = getCellState(key);
      
      // Check for existing slot indicators (same logic as in renderSlotComponentControls)
      const hasSwapSlot = swapSlotProp && cellState.properties ? cellState.properties[swapSlotProp] : null;
      const hasExistingSlot = cellState.slotComponentProps || cellState.statusIconText || hasSwapSlot;
      
      // If slot is not explicitly set but we have existing slot indicators, set it
      if (!cellState.hasOwnProperty('slot') && hasExistingSlot) {
        cellState.slot = true;
        console.log(`[initializeSlotVariables] Auto-set slot=true for cell ${key} based on existing indicators`);
      } else if (!cellState.hasOwnProperty('slot')) {
        // Initialize as false if no existing indicators
        cellState.slot = false;
        console.log(`[initializeSlotVariables] Initialized slot=false for cell ${key}`);
      }
    }
  }
  
  // Update the grid visuals to reflect the initialized slot states
  updateCellVisuals();
  console.log('[initializeSlotVariables] ✅ Slot variables initialized and grid updated');
}

// Removed test functions: discoverComponents, findOverflowComponent, discoverLibraryComponents

function analyzeSmartSlots(autoApply: boolean = false) {
  console.log('🧠 [UI] Analyzing table data for smart slots...');
  console.log('🧠 [UI] Current state:', {
    gridCols: state.gridCols,
    gridRows: state.gridRows,
    cellPropertiesSize: state.cellProperties.size
  });
  
  // Get current grid data and headers
  const headers: string[] = [];
  const gridData: string[][] = [];
  
  // Extract headers - check both old format (0-col) and new format (header-col)
  for (let col = 1; col <= state.gridCols; col++) {
    const oldKey = `0-${col - 1}`;
    const newKey = `header-${col}`;
    let cellProp = state.cellProperties.get(newKey) || state.cellProperties.get(oldKey);
    
    // Try to extract header text from properties
    let headerText = `Column ${col}`;
    if (cellProp?.properties) {
      // Look for Cell text property in properties object
      const cellTextProp = Object.keys(cellProp.properties).find(k => 
        k.toLowerCase().includes('text') && !k.toLowerCase().includes('second')
      );
      if (cellTextProp && cellProp.properties[cellTextProp]) {
        headerText = cellProp.properties[cellTextProp];
      }
    }
    headers.push(headerText);
    console.log(`🧠 [UI] Header ${col}: "${headerText}" (key: ${newKey})`);
  }
  
  // Extract data rows - use row,col format (1,1, 1,2, etc.)
  for (let row = 1; row <= state.gridRows; row++) {
    const rowData: string[] = [];
    for (let col = 1; col <= state.gridCols; col++) {
      const cellKey = `${row},${col}`;
      const cellProp = state.cellProperties.get(cellKey);
      
      // Try to extract cell text from properties
      let cellText = '';
      if (cellProp?.properties) {
        // Look for Cell text property in properties object - must be TEXT type, not BOOLEAN
        const cellTextProp = Object.keys(cellProp.properties).find(k => {
          const propValue = cellProp.properties[k];
          // Only return text properties (string values), not booleans
          return typeof propValue === 'string' && 
                 k.toLowerCase().includes('text') && 
                 !k.toLowerCase().includes('second') &&
                 !k.toLowerCase().includes('show');
        });
        if (cellTextProp && cellProp.properties[cellTextProp]) {
          cellText = String(cellProp.properties[cellTextProp]);
        }
      }
      rowData.push(cellText);
    }
    gridData.push(rowData);
    console.log(`🧠 [UI] Row ${row} data:`, rowData);
  }
  
  console.log(`📊 [UI] Analyzing ${state.gridCols} columns x ${state.gridRows} rows`);
  console.log(`📊 [UI] Headers:`, headers);
  console.log(`📊 [UI] Grid data:`, gridData);
  
  parent.postMessage({ 
    pluginMessage: { 
      type: 'analyze-smart-slots',
      gridData,
      headers,
      autoApply  // Pass autoApply flag to backend
    } 
  }, '*');
  
  console.log('✅ [UI] Message sent to backend');
  if (!autoApply) {
    showMessage('Analyzing table data for smart component suggestions...', 'success');
  }
}

// function collectCarbonKeys() {
//   console.log('🔑 Collecting Carbon component keys...');
//   parent.postMessage({ pluginMessage: { type: 'collect-carbon-keys' } }, '*');
//   showMessage('Collecting Carbon component key...', 'success');
// }

// Removed test functions: getComponentKeys, testSwapComponent

function createTable() {
  // Add loading state to button
  elements.createTableBtn.classList.add('loading');
  elements.createTableBtn.disabled = true;
  
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
  const includeToolbar = (document.getElementById('scanToolbarToggle') as HTMLInputElement)?.checked;
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
    includeToolbar,
    includeSelectable,
    includeExpandable
  };

  console.log(`🚀 Sending to Figma (${isUpdate ? 'Update' : 'Create'}):`, message);
  console.log(`🚀 cellProps keys:`, Object.keys(propsForFigma));
  
  // Debug: Check if "Show text" is properly set to false for slotted cells
  Object.keys(propsForFigma).forEach(key => {
    const cellData = propsForFigma[key];
    if (cellData.slotComponentProps) {
      console.log(`  📦 Cell ${key} has slot:`, cellData.properties);
    }
  });

  parent.postMessage({ pluginMessage: message }, '*');
}

function showLoader(message: string) {
  elements.loader.style.display = 'flex';
  elements.loaderText.textContent = message;
  // Add smooth fade-in animation
  requestAnimationFrame(() => {
    elements.loader.classList.add('show');
  });
}

function hideLoader() {
  elements.loader.classList.remove('show');
  setTimeout(() => {
    elements.loader.style.display = 'none';
    elements.loaderText.textContent = '';
  }, 250); // Match animation duration
}

function showMessage(text: string, type: 'success' | 'error') {
  elements.statusMessage.textContent = text;
  elements.statusMessage.className = `status ${type}`;
  elements.statusMessage.style.display = 'block';
  // Add smooth animation
  requestAnimationFrame(() => {
    elements.statusMessage.classList.add('show');
  });
  
  // Auto-hide success messages after 5 seconds
  if (type === 'success') {
    setTimeout(() => {
      hideMessage();
    }, 5000);
  }
}

function hideMessage() {
  elements.statusMessage.classList.remove('show');
  setTimeout(() => {
    elements.statusMessage.style.display = 'none';
  }, 250);
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
      console.log('[Tooltip] Cell properties for', cellKey, ':', cellProperties.properties);
      let tooltipContent = '<div class="tooltip-content">';

      // Add all properties to the tooltip with better formatting
      for (const [key, value] of Object.entries(cellProperties.properties)) {
        // Clean the property name for display
        const cleanKey = key.split('#')[0].trim();

        // Format value based on type with proper handling for all types
        let displayValue = '';
        
        // Special handling for "Swap slot" property: show component name instead of ID
        if (cleanKey.toLowerCase().includes('swap') && cleanKey.toLowerCase().includes('slot')) {
          displayValue = cellProperties.slotComponentName || String(value);
          console.log(`[Tooltip] Using component name for Swap slot: ${displayValue}`);
        } else if (typeof value === 'boolean') {
          displayValue = value ? 'true' : 'false';
        } else if (typeof value === 'number') {
          displayValue = String(value);
        } else if (typeof value === 'string') {
          displayValue = value || '(empty)';
        } else if (typeof value === 'object' && value !== null) {
          displayValue = JSON.stringify(value);
        } else if (value === null) {
          displayValue = 'null';
        } else if (value === undefined) {
          displayValue = 'undefined';
        } else {
          displayValue = String(value);
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
      
      // Skip creating slot checkboxes - handled by renderSlotComponentControls
      if (label.toLowerCase().includes('slot')) {
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

function renderSlotComponentControls(key: string, cellState: any) {
  const container = document.getElementById('dynamicPropertyFields');
  if (!container) {
    console.log(`[renderSlotComponentControls] Container not found`);
    return;
  }
  
  // Clear any existing slot controls first to prevent duplicates
  const existingSlotCheckbox = container.querySelector('#dynamicSlotCheckbox');
  if (existingSlotCheckbox) {
    existingSlotCheckbox.parentElement?.remove();
    console.log(`[renderSlotComponentControls] Removed existing slot checkbox for cell ${key}`);
  }
  
  const existingSlotSection = container.querySelector('.slot-component-section');
  if (existingSlotSection) {
    existingSlotSection.remove();
    console.log(`[renderSlotComponentControls] Removed existing slot section for cell ${key}`);
  }
  
  // Check if slot is enabled
  // First, try to find the swap slot property in availableProperties
  let swapSlotProp = state.selectedComponent?.availableProperties?.find((p: string) => 
    p.toLowerCase().includes('swap') && p.toLowerCase().includes('slot')
  );
  
  // If not found in availableProperties, check if it exists in cellState.properties
  if (!swapSlotProp && cellState && cellState.properties) {
    swapSlotProp = Object.keys(cellState.properties).find((p: string) => 
      p.toLowerCase().includes('swap') && p.toLowerCase().includes('slot')
    );
  }
  
  // Use the dedicated slot variable as the source of truth
  const currentSlotState = cellState.slot || false;
  
  // For backward compatibility, also check existing slot indicators
  const hasSwapSlot = swapSlotProp && cellState && cellState.properties ? cellState.properties[swapSlotProp] : null;
  const hasExistingSlot = cellState.slotComponentProps || cellState.statusIconText || hasSwapSlot;
  
  // If slot is not explicitly set but we have existing slot indicators, set it
  if (!cellState.hasOwnProperty('slot') && hasExistingSlot) {
    cellState.slot = true;
    console.log(`[renderSlotComponentControls] Auto-set slot=true for cell ${key} based on existing indicators`);
  }
  
  const hasSlot = cellState.slot || false;
  
  console.log(`[renderSlotComponentControls] Checking cell ${key}:`, {
    slot: cellState.slot,
    currentSlotState,
    hasSlotComponentProps: !!cellState.slotComponentProps,
    hasStatusIconText: !!cellState.statusIconText,
    hasSwapSlot: !!hasSwapSlot,
    swapSlotProp,
    swapSlotValue: (hasSwapSlot && swapSlotProp && cellState && cellState.properties) ? cellState.properties[swapSlotProp] : null,
    cellStateProperties: cellState && cellState.properties ? cellState.properties : {},
    hasSlot
  });
  
  // Always create slot checkbox for all cells, but only show slot properties if slot is enabled
  // This ensures consistent behavior across all cells
  
  // Determine the slot component type based on slotComponentProps
  const slotComponentType = cellState.slotComponentProps?.suggestedComponent || 
                           (cellState.slotComponentProps?.Status ? 'statusIcon' : null) ||
                           (cellState.slotComponentProps?.userName ? 'slotGroup' : null) ||
                           (cellState.slotComponentProps?.tagText ? 'tag' : null) ||
                           'unknown';
  
  console.log(`[renderSlotComponentControls] ✅ Rendering ${slotComponentType} controls for cell ${key}`);
  
  // Create slot component section
  const slotSection = document.createElement('div');
  slotSection.className = 'slot-component-section';
  slotSection.style.borderTop = '1px solid #e0e0e0';
  slotSection.style.marginTop = '16px';
  slotSection.style.paddingTop = '16px';
  
  // Only add slot properties if slot is enabled
  if (hasSlot) {
    // Section title based on component type
    const title = document.createElement('div');
    let sectionTitle = '🎨 Slot Component Properties';
    let componentName = 'Component';
  
  switch (slotComponentType) {
    case 'statusIcon':
      sectionTitle = '🎨 Status Icon Properties';
      componentName = 'Status Icon';
      break;
    case 'slotGroup':
      if (cellState.slotComponentProps?.userName) {
        sectionTitle = '👤 User Avatar Properties';
        componentName = 'Avatar + Text';
      } else if (cellState.slotComponentProps?.editKey) {
        sectionTitle = '⚡ Action Icons Properties';
        componentName = 'Edit + Delete Icons';
      } else {
        sectionTitle = '🎨 Slot Group Properties';
        componentName = 'Slot Group';
      }
      break;
    case 'tag':
      sectionTitle = '🏷️ Tag Properties';
      componentName = 'Tag Set';
      break;
    case 'overflow':
      sectionTitle = '📋 Overflow Properties';
      componentName = 'Overflow Menu';
      break;
    case 'edit':
      sectionTitle = '✏️ Edit Icon Properties';
      componentName = 'Edit Icon';
      break;
    case 'delete':
      sectionTitle = '🗑️ Delete Icon Properties';
      componentName = 'Delete Icon';
      break;
    default:
      sectionTitle = '🎨 Slot Component Properties';
      componentName = 'Component';
  }
  
  title.textContent = sectionTitle;
  title.style.fontWeight = '600';
  title.style.marginBottom = '12px';
  title.style.color = '#ffffff';
  slotSection.appendChild(title);
  
  // Component name (read-only - just shows which component is being used)
  if (swapSlotProp && cellState && cellState.properties && cellState.properties[swapSlotProp]) {
    const componentField = document.createElement('div');
    componentField.className = 'property-field';
    componentField.innerHTML = `
      <label>Component</label>
      <input type="text" class="styled-input component-name-readonly" value="${componentName}" readonly style="background: #1a1a1a; color: #8d8d8d; cursor: not-allowed; border: 1px solid #404040;">
    `;
    slotSection.appendChild(componentField);
  }
  
  // Show relevant fields based on component type
  if (slotComponentType === 'statusIcon') {
    // Status type dropdown
    const statusTypeField = document.createElement('div');
    statusTypeField.className = 'property-field';
    const statusTypeSelect = document.createElement('select');
    statusTypeSelect.className = 'styled-input';
    statusTypeSelect.id = 'status-icon-type';
    
    const statusTypes = ['Failed', 'Succeeded', 'Normal', 'In-progress', 'Caution major', 'Caution minor', 'Unknown', 'Undefined'];
    const currentType = cellState.statusIconType || 'Normal';
    
    statusTypes.forEach(type => {
      const option = document.createElement('option');
      option.value = type;
      option.textContent = type;
      if (type === currentType) option.selected = true;
      statusTypeSelect.appendChild(option);
    });
    
    statusTypeSelect.addEventListener('change', () => {
      const newType = statusTypeSelect.value;
      const textInput = document.getElementById('status-icon-text') as HTMLInputElement;
      const newText = textInput?.value || cellState.statusIconText || '';
      
      // Update cellState
      cellState.statusIconType = newType;
      cellState.statusIconText = newText;
      cellState.slotComponentProps = {
        'Status': newType,
        'Label': true,
        [`${newType} text`]: newText
      };
      
      console.log(`📝 Updated Status Icon: type="${newType}", text="${newText}"`);
    });
    
    const statusTypeLabel = document.createElement('label');
    statusTypeLabel.textContent = 'Status Type';
    statusTypeField.appendChild(statusTypeLabel);
    statusTypeField.appendChild(statusTypeSelect);
    slotSection.appendChild(statusTypeField);
    
    // Label text input
    const textField = document.createElement('div');
    textField.className = 'property-field';
    const textInput = document.createElement('input');
    textInput.type = 'text';
    textInput.className = 'styled-input';
    textInput.id = 'status-icon-text';
    textInput.value = cellState.statusIconText || '';
    textInput.placeholder = 'Enter label text...';
    
    textInput.addEventListener('input', () => {
      const newText = textInput.value;
      const typeSelect = document.getElementById('status-icon-type') as HTMLSelectElement;
      const newType = typeSelect?.value || cellState.statusIconType || 'Normal';
      
      // Update cellState
      cellState.statusIconText = newText;
      cellState.slotComponentProps = {
        'Status': newType,
        'Label': true,
        [`${newType} text`]: newText
      };
      
      console.log(`📝 Updated Status Icon text: "${newText}"`);
    });
    
    const textLabel = document.createElement('label');
    textLabel.textContent = 'Label Text';
    textField.appendChild(textLabel);
    textField.appendChild(textInput);
    slotSection.appendChild(textField);
    
  } else if (slotComponentType === 'slotGroup' && cellState.slotComponentProps?.userName) {
    // User name field (editable - user can change the name)
    const userNameField = document.createElement('div');
    userNameField.className = 'property-field';
    userNameField.innerHTML = `
      <label>User Name</label>
      <input type="text" id="slot-user-name" class="styled-input slot-editable-field" value="${cellState.slotComponentProps.userName}" style="background: #2a2a2a; color: #ffffff; border: 1px solid #404040;" placeholder="Enter user name">
    `;
    slotSection.appendChild(userNameField);
    
  } else if (slotComponentType === 'tag') {
    // Tag text field (editable - user can change the tags)
    const tagTextField = document.createElement('div');
    tagTextField.className = 'property-field';
    tagTextField.innerHTML = `
      <label>Tag Values</label>
      <input type="text" id="slot-tag-values" class="styled-input slot-editable-field" value="${cellState.slotComponentProps.tagText || ''}" style="background: #2a2a2a; color: #ffffff; border: 1px solid #404040;" placeholder="Enter tags (comma-separated)">
    `;
    slotSection.appendChild(tagTextField);
    
  } else if (slotComponentType === 'overflow') {
    // Overflow actions field (editable - user can see and edit actions)
    const overflowField = document.createElement('div');
    overflowField.className = 'property-field';
    const overflowActions = cellState.slotComponentProps?.overflowActions || 'View, Edit, Delete';
    overflowField.innerHTML = `
      <label>Actions (3+ actions)</label>
      <input type="text" id="slot-overflow-actions" class="styled-input slot-editable-field" value="${overflowActions}" style="background: #2a2a2a; color: #ffffff; border: 1px solid #404040;" placeholder="Enter actions (comma-separated)">
    `;
    slotSection.appendChild(overflowField);
    
  } else if (slotComponentType === 'edit' || slotComponentType === 'delete') {
    // Single action field (editable - user can see and change the action)
    const actionField = document.createElement('div');
    actionField.className = 'property-field';
    const actionValue = slotComponentType === 'edit' ? (cellState.slotComponentProps?.editAction || 'Edit') : (cellState.slotComponentProps?.deleteAction || 'Delete');
    actionField.innerHTML = `
      <label>Action</label>
      <input type="text" id="slot-single-action" class="styled-input slot-editable-field" value="${actionValue}" style="background: #2a2a2a; color: #ffffff; border: 1px solid #404040;" placeholder="Enter action name">
    `;
    slotSection.appendChild(actionField);
  }
  } // End of if (hasSlot) block
  
  container.appendChild(slotSection);
  
  // Create dynamic slot checkbox
  const slotCheckboxContainer = document.createElement('div');
  slotCheckboxContainer.className = 'property-field checkbox';
  
  const slotCheckbox = document.createElement('input');
  slotCheckbox.type = 'checkbox';
  slotCheckbox.id = 'dynamicSlotCheckbox';
  slotCheckbox.checked = cellState.slot || false;
  
  const slotLabel = document.createElement('label');
  slotLabel.textContent = 'Slot';
  slotLabel.setAttribute('for', 'dynamicSlotCheckbox');
  
  slotCheckboxContainer.appendChild(slotCheckbox);
  slotCheckboxContainer.appendChild(slotLabel);
  
  // Insert slot checkbox before the slot section
  container.insertBefore(slotCheckboxContainer, slotSection);
  
  // Set initial display state based on slot variable
  slotSection.style.display = cellState.slot ? 'block' : 'none';
  console.log(`[renderSlotComponentControls] ✅ Initial display set to: ${slotSection.style.display} (slot is ${cellState.slot ? 'true' : 'false'})`);
  
  // Add event listener for slot checkbox changes
  slotCheckbox.addEventListener('change', function(this: HTMLInputElement) {
    console.log(`[Slot Checkbox Event] 🎯 Change event fired! Checkbox is now: ${this.checked ? 'CHECKED' : 'UNCHECKED'}`);
    
    // Only update UI visibility - don't modify cellState until Save is clicked
    slotSection.style.display = this.checked ? 'block' : 'none';
    console.log(`[Slot Checkbox Event] ✅ Set slotSection display to: ${slotSection.style.display} (state will be saved on Save click)`);
  });
  
  console.log(`[renderSlotComponentControls] ✅ Dynamic slot checkbox created with state: ${cellState.slot}`);
  
  // Add change event listeners to editable slot fields (exclude read-only component name)
  const editableInputs = slotSection.querySelectorAll('.slot-editable-field');
  console.log(`[renderSlotComponentControls] 🔍 Found ${editableInputs.length} editable slot fields`);
  
  editableInputs.forEach((input: Element) => {
    const inputElement = input as HTMLInputElement;
    console.log(`[renderSlotComponentControls] 🔗 Adding listener to field:`, inputElement.id);
    
    // Remove immediate change listener - changes should only be applied on Save
    // This prevents changes from being saved when user types but doesn't click Save
  });
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

      // Add change listener for fields that other fields depend on
      // This re-renders the form to show/hide dependent fields
      const hasDependent = model.some(f => f.dependsOn === name);
      if (hasDependent) {
        select.addEventListener('change', () => {
          console.log(`[DEBUG] ${name} changed to ${select.value}, triggering re-render for dependent fields`);
          // Update the props object with the new value
          props[name] = select.value;
          // Re-render the form to show/hide dependent fields
          renderDynamicPropertyFieldsFromModel(model, props);
        });
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
    const slotField = document.querySelector('.property-field.checkbox') as HTMLElement;
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
  // Set custom cell text toggle based on saved state
  if (elements.customCellTextToggle) {
    // Restore the custom cell text toggle state from cell properties
    const cellState = state.cellProperties.get(state.currentEditingCell || '');
    const savedToggleState = cellState?.customCellTextEnabled;
    elements.customCellTextToggle.checked = savedToggleState !== undefined ? savedToggleState : true; // Default to true if not set
    // Ensure the input is shown when toggle is checked
    if (elements.customCellTextContainer) {
      elements.customCellTextContainer.style.display = cellTextPropKey ? 'none' : (elements.customCellTextToggle.checked ? 'block' : 'none');
      console.log('[DEBUG] Setting customCellTextContainer display to', cellTextPropKey ? 'none' : (elements.customCellTextToggle.checked ? 'block' : 'none'), 'in updateStaticFieldsVisibilityAndValues');
    }
  }
  console.log('[DEBUG] updateStaticFieldsVisibilityAndValues elements presence', {
    customCellText: !!elements.customCellText,
    secondCellText: !!elements.secondCellText,
    secondLineContainer: !!elements.secondLineContainer,
    secondCellTextContainer: !!elements.secondCellTextContainer,
    // slotCheckbox: !!elements.slotCheckbox, // Removed - using dynamic slot checkbox only
    state: !!elements.state
  });
  if (elements.secondLineContainer && elements.secondCellText && elements.customCellTextToggle) {
    elements.secondLineContainer.style.display = secondTextPropKey ? 'none' : (!!elements.secondCellText.value && elements.customCellTextToggle?.checked ? '' : 'none');
  }
  if (elements.secondCellTextContainer && elements.secondCellText && elements.customCellTextToggle) {
    elements.secondCellTextContainer.style.display = secondTextPropKey ? 'none' : (!!elements.secondCellText.value && elements.customCellTextToggle?.checked ? 'block' : 'none');
  }
  const slotField = document.querySelector('.property-field.checkbox') as HTMLElement;
  if (slotField) slotField.style.display = slotPropKey ? 'none' : '';
  const stateField = elements.state?.parentElement;
  if (stateField) stateField.style.display = statePropKey ? 'none' : '';

  // Populate static fields from props or reset to default
  if (elements.showText) {
    elements.showText.checked = showTextPropKey ? (props[showTextPropKey] || false) : true; // Always default to true since checkbox is hidden
  }
  if (elements.customCellTextToggle) {
    // Restore the custom cell text toggle state from cell properties
    const cellState = state.cellProperties.get(state.currentEditingCell || '');
    const savedToggleState = cellState?.customCellTextEnabled;
    elements.customCellTextToggle.checked = savedToggleState !== undefined ? savedToggleState : true; // Default to true if not set
    console.log('[DEBUG] Restored customCellTextToggle state:', elements.customCellTextToggle.checked, 'for cell:', state.currentEditingCell);
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
  // Slot checkbox is now handled dynamically by renderSlotComponentControls
  console.log('[DEBUG] Slot checkbox state managed by renderSlotComponentControls');
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
  // Allow property editor to open regardless of mode
  // if (state.mode !== 'edit') return;

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

    // Populate selectedCells based on current apply mode
    // This ensures the property split logic works correctly when saving
    state.selectedCells.clear();
    if (!key.startsWith('header-') && key !== 'footer') {
      // For body cells, populate selectedCells based on apply mode
      const [row, col] = key.split(',').map(Number);
      const applyMode = state.applyMode as 'cell' | 'row' | 'column'; // Type assertion for user-changeable mode
      
      switch (applyMode) {
        case 'cell':
          // Single cell only
          state.selectedCells.add(key);
          console.log(`[openPropertyEditorInternal] Apply mode: cell - selected ${key}`);
          break;
        case 'row':
          // All cells in the row
          for (let c = 1; c <= state.gridCols; c++) {
            const cellKey = `${row},${c}`;
            state.selectedCells.add(cellKey);
          }
          console.log(`[openPropertyEditorInternal] Apply mode: row - selected ${state.selectedCells.size} cells in row ${row}`);
          break;
        case 'column':
          // All cells in the column
          for (let r = 1; r <= state.gridRows; r++) {
            const cellKey = `${r},${col}`;
            state.selectedCells.add(cellKey);
          }
          console.log(`[openPropertyEditorInternal] Apply mode: column - selected ${state.selectedCells.size} cells in column ${col}`);
          break;
      }
    } else {
      // For header/footer, always single cell
      state.selectedCells.add(key);
      console.log(`[openPropertyEditorInternal] Header/footer cell - selected ${key}`);
    }

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
      const cellState = getCellState(key);
      props = cellState.properties || {};

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
      const cellState = getCellState(key);
      props = cellState.properties || {};
      label = 'Footer';
      console.log('[DEBUG] openPropertyEditor footer', { props });

      // For footer, only show the "Type" property
      const footerModel = FOOTER_MODEL.filter(field => field.label === "Type");
      renderDynamicPropertyFieldsFromModel(footerModel, props);
    } else {
      console.log(`[DEBUG] openPropertyEditor body cell ${key} - state.cellProperties exists:`, !!state.cellProperties);
      console.log(`[DEBUG] openPropertyEditor body cell ${key} - state.cellProperties type:`, typeof state.cellProperties);
      
      const cellState = getCellState(key);
      console.log(`[DEBUG] openPropertyEditor body cell ${key} - getCellState returned:`, cellState);
      console.log(`[DEBUG] openPropertyEditor body cell ${key} - cellState.properties exists:`, !!cellState?.properties);
      
      if (!cellState) {
        console.error(`[DEBUG] openPropertyEditor body cell ${key} - getCellState returned undefined!`);
        return;
      }
      
      props = cellState.properties || {};
      const [row, col] = key.split(',');
      label = `(${row},${col})`;
      console.log('[DEBUG] openPropertyEditor body', { availableProps, propertyTypes, props });
      renderDynamicPropertyFields(availableProps, propertyTypes, props);
      
      // Add slot component controls if slot is enabled
      renderSlotComponentControls(key, cellState);
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
        const col = key.split('-')[1];
        width = getDefaultColumnWidthUI(parseInt(col));
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
        const col = key.split(',')[1];
        width = getDefaultColumnWidthUI(parseInt(col));
      }
      console.log('[DEBUG] Body col width UI exists?', !!elements.colWidthContainer, !!elements.colWidthInput, 'value to set:', width, 'for key', key);
      if (elements.colWidthContainer) elements.colWidthContainer.style.display = 'none';
      if (elements.colWidthInput) elements.colWidthInput.value = String(width);
      console.log(`[DEBUG] Body cell ${key} - colWidth: ${width}, container display: ${elements.colWidthContainer.style.display}`);
    }

    // Capture original form values before showing the editor
    originalFormValues = captureOriginalFormValues();
    
    // Set title and show editor
    elements.propertyEditorTitle.innerHTML = `Cell Properties <span>${label}</span>`;
    elements.editingCellCoords.textContent = label;
    elements.propertyEditorOverlay.style.display = 'block';
    elements.propertyEditor.style.display = 'block';
    
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

// Helper function to determine if a property is structural (should apply to all cells in column/row)
// or content (should apply only to current cell)
function isStructuralProperty(propName: string, propType: string): boolean {
  const propLower = propName.toLowerCase();
  
  // IMPORTANT: Check propType FIRST before checking name
  // This ensures BOOLEAN checkboxes are always structural, even if they have "text" in the name
  
  // INCLUDE: Boolean properties (checkboxes) are ALWAYS structural
  if (propType === 'BOOLEAN') {
    console.log(`  ⚙️ Property "${propName}" is BOOLEAN → STRUCTURAL property`);
    return true;
  }
  
  // INCLUDE: Variant properties (dropdowns) are ALWAYS structural
  if (propType === 'VARIANT') {
    console.log(`  ⚙️ Property "${propName}" is VARIANT → STRUCTURAL property`);
    return true;
  }
  
  // INCLUDE: Component swap properties are structural
  if (propLower.includes('swap')) {
    console.log(`  ⚙️ Property "${propName}" is swap property → STRUCTURAL property`);
    return true;
  }
  
  // EXCLUDE: TEXT type properties are content (actual text values)
  if (propType === 'TEXT') {
    console.log(`  📝 Property "${propName}" is TEXT → CONTENT property`);
    return false;
  }
  
  // EXCLUDE: Properties with "text" in name that are not BOOLEAN/VARIANT (e.g., text inputs)
  // This catches properties like "Cell text", "Second cell text" but NOT "Second text line" (which is BOOLEAN)
  if (propLower.includes('text')) {
    console.log(`  📝 Property "${propName}" contains 'text' and is not BOOLEAN/VARIANT → CONTENT property`);
    return false;
  }
  
  // Default: structural (safer to apply to all cells)
  console.log(`  ⚙️ Property "${propName}" (type: ${propType}) → STRUCTURAL property (default)`);
  return true;
}

// Filter properties to get only structural properties
function filterStructuralProperties(allProps: any, propertyTypes: { [key: string]: any }): any {
  console.log(`[filterStructuralProperties] Filtering from ${Object.keys(allProps).length} properties`);
  console.log(`[filterStructuralProperties] All property types:`, propertyTypes);
  const structural: any = {};
  
  for (const [propName, value] of Object.entries(allProps)) {
    const propType = propertyTypes[propName] || 'UNKNOWN';
    console.log(`[filterStructuralProperties] Checking "${propName}" (type: ${propType}, value: ${value})`);
    
    if (isStructuralProperty(propName, propType)) {
      structural[propName] = value;
      console.log(`  ✅ Including "${propName}" = ${value}`);
    } else {
      console.log(`  ❌ Excluding "${propName}" (content property)`);
    }
  }
  
  // Always include the main slot boolean if it exists (for backward compatibility)
  if (allProps.hasOwnProperty('slot')) {
    structural['slot'] = allProps['slot'];
    console.log(`  ✅ Including "slot" = ${allProps['slot']} (special case - dedicated slot variable)`);
  }
  
  console.log(`[filterStructuralProperties] Result: ${Object.keys(structural).length} structural properties:`, Object.keys(structural));
  return structural;
}

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
          // Always save the input value, even if it's empty (to allow clearing text)
          newProps[name] = inputValue || defaultValue || '';
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
      
      // Save the custom cell text toggle state in cell properties
      const cellState = getCellState(key);
      cellState.customCellTextEnabled = showText;

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

      // Save the dedicated slot variable
      const dynamicSlotCheckbox = document.getElementById('dynamicSlotCheckbox') as HTMLInputElement;
      if (dynamicSlotCheckbox) {
        cellState.slot = dynamicSlotCheckbox.checked;
        console.log(`🔍 UI: Saved cellState.slot = ${cellState.slot}`);
      }
      
      // Also save as a property - ALWAYS update it to match cellState.slot
      const slotBooleanProp = availableProps.find(p => p.toLowerCase().includes('slot') && propertyTypes[p] === 'BOOLEAN') || 'Slot';
      props[slotBooleanProp] = cellState.slot || false;
      console.log(`🔍 UI: Saving Slot property: ${slotBooleanProp} = ${props[slotBooleanProp]} (synced with cellState.slot)`);


      // Save dynamic properties
      console.log(`[saveCellProperties] Processing ${availableProps.length} available properties for ${key}`);
      for (const propName of availableProps) {
        const type = propertyTypes[propName];
        const input = document.getElementById(`dynamic-${propName}`) as HTMLInputElement | HTMLSelectElement;
        
        if (!input) {
          // If no input exists, check if we already set this property above
          // (e.g., custom cell text was already saved from elements.customCellText)
          if (props[propName] !== undefined) {
            console.log(`✅ UI: Property already set: ${propName} = ${props[propName]}`);
            continue;
          }
          
          // Otherwise, preserve the existing value from cellState
          // This is important for properties that don't have UI inputs (like swap slot)
          const cellState = getCellState(key);
          if (cellState.properties && cellState.properties[propName] !== undefined) {
            props[propName] = cellState.properties[propName];
            console.log(`🔄 UI: Preserving existing property (no input): ${propName} = ${props[propName]}`);
          } else {
            console.log(`⚠️ UI: No input and no existing value for: ${propName}`);
          }
          continue;
        }
        
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
      console.log(`[saveCellProperties] Final props to save for ${key}:`, Object.keys(props));

      // Collect and save slot component properties from form fields
      const currentCellState = getCellState(key);
      
      // Check if slot is enabled from the props we just saved (NOT from DOM element which may be stale)
      const slotPropName = availableProps.find(p => p.toLowerCase().includes('slot') && propertyTypes[p] === 'BOOLEAN');
      const isSlotEnabled = slotPropName ? (props[slotPropName] === true || props[slotPropName] === 'true') : false;
      console.log(`[saveCellProperties] Slot enabled check: ${slotPropName || 'none'} = ${slotPropName ? props[slotPropName] : 'N/A'} → isSlotEnabled = ${isSlotEnabled}`);
      
      if (isSlotEnabled && currentCellState.slotComponentProps) {
        console.log(`[saveCellProperties] Slot is enabled, collecting slot properties for ${key}`);
        
        // Collect current values from slot form fields
        const userNameField = document.getElementById('slot-user-name') as HTMLInputElement;
        const tagTextField = document.getElementById('slot-tag-values') as HTMLInputElement;
        const actionsField = document.getElementById('slot-overflow-actions') as HTMLInputElement;
        const editField = document.getElementById('slot-single-action') as HTMLInputElement;
        
        // Update slotComponentProps with current form values
        if (userNameField && userNameField.value) {
          currentCellState.slotComponentProps.userName = userNameField.value;
          console.log(`[saveCellProperties] Updated userName: "${userNameField.value}"`);
        }
        
        if (tagTextField && tagTextField.value) {
          currentCellState.slotComponentProps.tagText = tagTextField.value;
          console.log(`[saveCellProperties] Updated tagText: "${tagTextField.value}"`);
        }
        
        if (actionsField && actionsField.value) {
          currentCellState.slotComponentProps.overflowActions = actionsField.value;
          console.log(`[saveCellProperties] Updated overflowActions: "${actionsField.value}"`);
        }
        
        if (editField && editField.value) {
          // Determine if this is edit or delete action based on the slot component type
          const slotComponentType = currentCellState.slotComponentProps?.suggestedComponent || 'unknown';
          if (slotComponentType === 'edit') {
            currentCellState.slotComponentProps.editAction = editField.value;
            console.log(`[saveCellProperties] Updated editAction: "${editField.value}"`);
          } else if (slotComponentType === 'delete') {
            currentCellState.slotComponentProps.deleteAction = editField.value;
            console.log(`[saveCellProperties] Updated deleteAction: "${editField.value}"`);
          }
        }
        
        console.log(`[saveCellProperties] Final slotComponentProps for ${key}:`, currentCellState.slotComponentProps);
      } else if (!isSlotEnabled && currentCellState.slotComponentProps) {
        // If slot is disabled but slotComponentProps exists, we keep it for future use
        // but don't display it (handled by hasSlotEnabled check in display logic)
        console.log(`[saveCellProperties] Slot is disabled for ${key}, slotComponentProps preserved but not displayed`);
      }

      // Col width logic: always set colWidth as a top-level property on first row of column
      let colWidth: number | undefined = undefined;
      if (elements.colWidthInput && elements.colWidthInput.value) {
        colWidth = parseInt(elements.colWidthInput.value, 10);
      } else {
        console.log('[DEBUG] colWidthInput missing or empty when saving cell (body)');
      }
      const applyProps = (key: string, props: any, colWidthTopLevel?: number, structuralOnly: boolean = false) => {
        const cellState = getCellState(key);
        
        // Check if this is the cell currently being edited
        const isCurrentEditingCell = (key === state.currentEditingCell);
        
        console.log(`[applyProps] Applying to ${key}, structuralOnly: ${structuralOnly}, isCurrentEditingCell: ${isCurrentEditingCell}`);
        
        // Save special properties BEFORE updating
        const savedIsCheckbox = cellState.isCheckbox;
        const savedSlotComponentProps = cellState.slotComponentProps;
        const savedStatusIconText = cellState.statusIconText;
        const savedStatusIconType = cellState.statusIconType;
        
        // Save existing text properties if we're applying structural only
        const savedTextProperties: any = {};
        if (structuralOnly && cellState.properties) {
          for (const [propName, value] of Object.entries(cellState.properties)) {
            const propType = propertyTypes[propName] || 'UNKNOWN';
            if (!isStructuralProperty(propName, propType)) {
              savedTextProperties[propName] = value;
              console.log(`  💾 Saving text property for preservation: "${propName}" = ${value}`);
            }
          }
        }
        
        // Also save slot-related properties from the old properties
        // BUT only for cells that are NOT the current editing cell
        const oldSlotProp = availableProps.find(p => 
          p.toLowerCase().includes('slot') && propertyTypes[p] === 'BOOLEAN');
        const savedSlotValue = (!isCurrentEditingCell && oldSlotProp && cellState.properties) ? cellState.properties[oldSlotProp] : undefined;
        
        // Now update the properties
        cellState.properties = props;
        
        // Restore text properties if we applied structural only
        if (structuralOnly && Object.keys(savedTextProperties).length > 0) {
          for (const [propName, value] of Object.entries(savedTextProperties)) {
            cellState.properties[propName] = value;
            console.log(`  ✅ Restored text property: "${propName}" = ${value}`);
          }
        }
        
        // CRITICAL: Sync the dedicated cellState.slot variable with the Slot property
        // This ensures both the property AND the dedicated variable are in sync
        if (oldSlotProp && cellState.properties[oldSlotProp] !== undefined) {
          cellState.slot = cellState.properties[oldSlotProp] === true || cellState.properties[oldSlotProp] === 'true';
          console.log(`  🔄 Synced cellState.slot = ${cellState.slot} from property ${oldSlotProp} = ${cellState.properties[oldSlotProp]} for ${key}`);
        }
        
        // Restore special properties that should not be overwritten
        // For the current editing cell, we want to use the NEW values from props
        if (savedIsCheckbox !== undefined && !isCurrentEditingCell) {
          cellState.isCheckbox = savedIsCheckbox;
        }
        if (savedSlotComponentProps !== undefined && !isCurrentEditingCell) {
          cellState.slotComponentProps = savedSlotComponentProps;
          console.log(`  ✅ Preserved slotComponentProps for ${key}:`, savedSlotComponentProps);
          
          // ✅ FIX: When applying structural properties (which includes slot boolean),
          // DO NOT restore the old slot boolean value - use the NEW value from props instead
          // This allows unchecking slot in "apply to column" mode to work correctly
          if (oldSlotProp && savedSlotValue !== undefined && structuralOnly === false) {
            // structuralOnly === false means we're NOT applying structural props,
            // so we should preserve the old slot value
            cellState.properties[oldSlotProp] = savedSlotValue;
            // Re-sync cellState.slot after restoring the property
            cellState.slot = savedSlotValue === true || savedSlotValue === 'true';
            console.log(`  ✅ Preserved slot boolean property ${oldSlotProp} = ${savedSlotValue} for ${key} (cellState.slot = ${cellState.slot})`);
          }
          // If structuralOnly === true, the slot boolean was already applied from props,
          // so we don't need to restore the old value
        } else if (isCurrentEditingCell) {
          // For the current editing cell, preserve slotComponentProps but NOT the slot boolean
          // The slot boolean value in props is the new value the user just set
          if (savedSlotComponentProps !== undefined) {
            cellState.slotComponentProps = savedSlotComponentProps;
            console.log(`  🔄 Current editing cell ${key}: preserved slotComponentProps but using new slot boolean from props`);
          }
        }
        if (savedStatusIconText !== undefined && !isCurrentEditingCell) {
          cellState.statusIconText = savedStatusIconText;
          console.log(`  ✅ Preserved statusIconText for ${key}:`, savedStatusIconText);
        }
        if (savedStatusIconType !== undefined && !isCurrentEditingCell) {
          cellState.statusIconType = savedStatusIconType;
          console.log(`  ✅ Preserved statusIconType for ${key}:`, savedStatusIconType);
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
        // Single cell: apply all properties (both structural and content)
        if (colWidth && row === 1) {
          applyProps(state.currentEditingCell, props, colWidth, false);
        } else {
          applyProps(state.currentEditingCell, props, undefined, false);
        }
      } else if (state.applyMode === 'row') {
        console.log(`[saveCellProperties] Applying to ROW ${row} - splitting properties`);
        
        // Get structural properties only
        const structuralProps = filterStructuralProperties(props, propertyTypes);
        
        for (let c = 1; c <= state.gridCols; c++) {
          const k = `${row},${c}`;
          if (state.selectedCells.has(k)) {
            if (k === state.currentEditingCell) {
              // Current cell: apply ALL properties (structural + content)
              console.log(`  📍 Current cell ${k}: applying ALL properties`);
              if (colWidth && row === 1) {
                applyProps(k, { ...props }, colWidth, false);
              } else {
                applyProps(k, { ...props }, undefined, false);
              }
            } else {
              // Other cells in row: apply ONLY structural properties
              console.log(`  🔧 Other cell ${k}: applying STRUCTURAL properties only`);
              if (colWidth && row === 1) {
                applyProps(k, { ...structuralProps }, colWidth, true);
              } else {
                applyProps(k, { ...structuralProps }, undefined, true);
              }
            }
            const cellState = getCellState(k);
            if (cellState.isCheckbox !== undefined) {
              cellState.isCheckbox = cellState.isCheckbox;
            }
          }
        }
      } else if (state.applyMode === 'column') {
        console.log(`[saveCellProperties] Applying to COLUMN ${col} - splitting properties`);
        
        // Get structural properties only
        const structuralProps = filterStructuralProperties(props, propertyTypes);
        
        for (let r = 1; r <= state.gridRows; r++) {
          const k = `${r},${col}`;
          if (state.selectedCells.has(k)) {
            if (k === state.currentEditingCell) {
              // Current cell: apply ALL properties (structural + content)
              console.log(`  📍 Current cell ${k}: applying ALL properties`);
              if (colWidth && r === 1) {
                applyProps(k, { ...props }, colWidth, false);
              } else {
                applyProps(k, { ...props }, undefined, false);
              }
            } else {
              // Other cells in column: apply ONLY structural properties
              console.log(`  🔧 Other cell ${k}: applying STRUCTURAL properties only`);
              if (colWidth && r === 1) {
                applyProps(k, { ...structuralProps }, colWidth, true);
              } else {
                applyProps(k, { ...structuralProps }, undefined, true);
              }
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
      const aiChoice = (document.getElementById('aiSource') as HTMLSelectElement | null)?.value || 'watsonx';
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
        
        // Show loader for Faker generation
        showLoader('Generating sample data...');
        console.log('[saveCellProperties] Showing loader for Faker generation');
        
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
        // Analytics: track watsonx generation intent
        try {
          const colCtx = (state.applyMode === 'column' && state.currentEditingCell) ? { column: Number(state.currentEditingCell.split(',')[1]) } : {};
          trackEvent('ai_generate_watsonx', { mode, ...colCtx, prompt, length: (prompt || '').length });
        } catch (e) { console.warn('[analytics] track error', e); }
        const endpoint = 'https://us-south.ml.cloud.ibm.com';
        const apiKey = ''; // Not needed when using proxy server
        const useAccessToken = false;
        const accessToken = '';
        // API key validation removed - now using Code Engine proxy server
        const remember = true;
        const useProxy = true;
        const proxyUrl = 'https://application-e9.21hwt6k1vujm.us-east.codeengine.appdomain.cloud';
        
        // Show loader for watsonx.ai generation
        showLoader('Generating content with AI...');
        console.log('[saveCellProperties] Showing loader for watsonx.ai generation');
        
        pendingFakerContext = { mode, key, fakerMethod: 'watsonx' };
        parent.postMessage({ pluginMessage: { type: 'generate-watsonx-data', prompt, endpoint, apiKey, accessToken, useAccessToken, useProxy, proxyUrl, count, remember } }, '*');
        return;
      }
    }
    
    // For generated tables, save cell properties to backend without rebuilding table
    if (state.tableFrameId) {
      console.log('💾 [UI] Saving cell properties to backend...');
      
      // Apply sorting before saving to ensure sorted state is preserved
      if (state.gridRows > 0 && state.gridCols > 0) {
        const sortedCellProperties = sortColumnData(state.cellProperties, state.gridCols, state.gridRows);
        state.cellProperties = sortedCellProperties;
        console.log('🔄 [UI] Applied sorting before saving cell properties');
      }
      
      // Convert Map to object properly
      const cellPropertiesObj: { [key: string]: any } = {};
      for (const [key, value] of state.cellProperties.entries()) {
        cellPropertiesObj[key] = value;
      }
      
      // Send cell properties to backend for saving (without rebuilding table)
      parent.postMessage({
        pluginMessage: {
          type: 'save-cell-properties',
          tableId: state.tableFrameId,
          cellProperties: cellPropertiesObj
        }
      }, '*');
      console.log('💾 [UI] Sent save-cell-properties message to backend');
    }
    
    console.log(`[DEBUG] saveCellProperties - Before closePropertyEditor: gridCols=${state.gridCols}, gridRows=${state.gridRows}`);
    console.log(`[DEBUG] saveCellProperties - cellProperties keys before close:`, Array.from(state.cellProperties.keys()));
    
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
  console.log(`[DEBUG] saveCellProperties - Before createGrid: gridCols=${state.gridCols}, gridRows=${state.gridRows}`);
  console.log(`[DEBUG] saveCellProperties - cellProperties keys before createGrid:`, Array.from(state.cellProperties.keys()));
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
  // If we don't have component info yet, request it and return
  if (!state.selectedComponent && state.tableFrameId) {
    console.log(`[updateCellVisuals] No component info available, requesting...`);
    parent.postMessage({
      pluginMessage: {
        type: 'request-component-info',
        tableId: state.tableFrameId
      }
    }, '*');
    return;
  }
  
  const availableProps: string[] = state.selectedComponent?.availableProperties || [];
  console.log(`[updateCellVisuals] availableProps:`, availableProps);
  console.log(`[updateCellVisuals] selectedComponent:`, state.selectedComponent);

  // Update body cells
  document.querySelectorAll('.cell').forEach(cell => {
    const divCell = cell as HTMLDivElement;
    const key = `${divCell.dataset.row},${divCell.dataset.col}`;
    const props = getCellState(key);
    console.log(`[updateCellVisuals] Cell ${key}:`, props);

    if (props && props.properties && Object.keys(props.properties).length > 0) {
      let displayText = '';
      let hasSlotEnabled = false;

      // If we have selectedComponent, use its property types
      if (state.selectedComponent && availableProps.length > 0) {
        // Check for slot using the actual component property as source of truth
        const slotPropName = availableProps.find(p => 
          state.selectedComponent!.propertyTypes[p] === 'BOOLEAN' && 
          p.toLowerCase().includes('slot') && 
          !p.toLowerCase().includes('swap')
        );
        
        if (slotPropName && props.properties[slotPropName] !== undefined) {
          hasSlotEnabled = props.properties[slotPropName] === true;
          console.log(`[updateCellVisuals] Cell ${key} slot state: ${hasSlotEnabled} (from property ${slotPropName} = ${props.properties[slotPropName]})`);
        } else {
          // Fallback to dedicated slot variable if property doesn't exist
          hasSlotEnabled = props.slot || false;
          console.log(`[updateCellVisuals] Cell ${key} slot state: ${hasSlotEnabled} (fallback from cellState.slot)`);
        }
        
        // Check if "Show text" is disabled
        const showTextProp = availableProps.find(p =>
          state.selectedComponent!.propertyTypes[p] === 'BOOLEAN' && 
          p.toLowerCase().includes('show') && 
          p.toLowerCase().includes('text'));
        const showTextEnabled = !showTextProp || props.properties[showTextProp] !== false;

        // Always show cell text in preview grid (even when Show text is disabled for smart slots)
        // This helps users see the actual data values while smart slots are applied
        if (true) { // Always show text in preview grid
        
        // Only use slot component data if slot is actually enabled
        if (hasSlotEnabled && props.slotComponentProps) {
          console.log(`[updateCellVisuals] Cell ${key} has slot enabled and slotComponentProps:`, props.slotComponentProps);
          if (props.slotComponentProps.userName) {
            displayText = props.slotComponentProps.userName;
            console.log(`[updateCellVisuals] ✅ Using userName for display: "${displayText}"`);
          } else if (props.slotComponentProps.tagText) {
            displayText = props.slotComponentProps.tagText;
            console.log(`[updateCellVisuals] ✅ Using tagText for display: "${displayText}"`);
          } else if (props.slotComponentProps.overflowActions) {
            displayText = props.slotComponentProps.overflowActions;
            console.log(`[updateCellVisuals] ✅ Using overflowActions for display: "${displayText}"`);
          } else if (props.slotComponentProps.editAction) {
            displayText = props.slotComponentProps.editAction;
            console.log(`[updateCellVisuals] ✅ Using editAction for display: "${displayText}"`);
          } else if (props.slotComponentProps.deleteAction) {
            displayText = props.slotComponentProps.deleteAction;
            console.log(`[updateCellVisuals] ✅ Using deleteAction for display: "${displayText}"`);
          }
        }
        
        // If slot is disabled or no slot component data, use regular text properties
        if (!displayText) {
        for (const propName of availableProps) {
          const type = state.selectedComponent.propertyTypes[propName];
          if (type === 'TEXT' && typeof props.properties[propName] === 'string' && props.properties[propName].toString().trim() !== '') {
            displayText = props.properties[propName].toString();
            break; // Use the first TEXT property found
          }
        }
        }
        // Fallbacks when the detected TEXT key doesn't exist in props
        if (!displayText) {
          // 1) Prefer plain 'Cell text' if present
          if (typeof props.properties['Cell text'] === 'string' && props.properties['Cell text'].toString().trim() !== '') {
            displayText = props.properties['Cell text'].toString();
          } else {
            // 2) Any key containing 'text' (includes hashed variants)
            const textProps = Object.keys(props.properties).filter(prop =>
              typeof props.properties[prop] === 'string' &&
              props.properties[prop].toString().trim() !== '' &&
              prop.toLowerCase().includes('text') &&
              !prop.toLowerCase().includes('second')
            );
            if (textProps.length > 0) {
              // Prefer exact 'Cell text' when listed among matches
              const preferred = textProps.find(p => p === 'Cell text') || textProps[0];
              displayText = props.properties[preferred].toString();
            } else {
              // 3) Last-resort: any non-empty string property
              const anyStringProp = Object.keys(props.properties).find(prop =>
                typeof props.properties[prop] === 'string' && props.properties[prop].toString().trim() !== ''
              );
              if (anyStringProp) {
                displayText = props.properties[anyStringProp].toString();
              }
            }
          }
        }
        } // End of showTextEnabled check
        // If no text and slot is checked, show empty
        if (!displayText && hasSlotEnabled) {
          displayText = '';
        }
      } else {
        // Prefer plain 'Cell text' when component props are unavailable
        if (typeof props.properties['Cell text'] === 'string' && props.properties['Cell text'].toString().trim() !== '') {
          displayText = props.properties['Cell text'].toString();
        } else {
          // Fallback: look for common text property names in the loaded properties
          const textProps = Object.keys(props.properties).filter(prop =>
            prop.toLowerCase().includes('text') &&
            !prop.toLowerCase().includes('second') &&
            typeof props.properties[prop] === 'string' &&
            props.properties[prop].toString().trim() !== ''
          );

          if (textProps.length > 0) {
            // Prefer the un-hashed key 'Cell text' over hashed ones if both exist
            const preferred = textProps.find(p => p === 'Cell text') || textProps[0];
            displayText = props.properties[preferred].toString();
          }
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
        const props = getCellState(key);
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
      const props = getCellState('footer');
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
  // Button should be enabled when:
  // 1. sizeConfirmed is true (user has confirmed table size)
  // 2. OR hasComponent is true AND grid has cells with data
  const hasGridData = state.gridCols > 0 && state.gridRows > 0;
  const shouldEnable = state.sizeConfirmed || (state.hasComponent && hasGridData);
  
  elements.createTableBtn.disabled = !shouldEnable;
  elements.createTableBtn.style.opacity = shouldEnable ? '1' : '0.5';
  elements.createTableBtn.style.cursor = shouldEnable ? 'pointer' : 'not-allowed';
  
  // Remove loading class if button is enabled (safeguard against stuck loading state)
  if (shouldEnable && elements.createTableBtn.classList.contains('loading')) {
    elements.createTableBtn.classList.remove('loading');
    console.log(`[updateCreateButtonState] Removed stuck loading class from button`);
  }
  
  // Ensure button always has text - prevent empty button
  if (!elements.createTableBtn.textContent || elements.createTableBtn.textContent.trim() === '') {
    elements.createTableBtn.textContent = state.tableFrameId ? 'Update Table' : 'Create Table';
    console.log(`[updateCreateButtonState] Button text was empty, set to: ${elements.createTableBtn.textContent}`);
  }
  
  // Update title based on state
  if (!state.hasComponent) {
    elements.createTableBtn.title = 'Select a component first';
  } else if (!hasGridData) {
    elements.createTableBtn.title = 'Configure table size and properties';
  } else {
    elements.createTableBtn.title = state.tableFrameId ? 'Update table in Figma' : 'Create table in Figma';
  }
  
  console.log(`[updateCreateButtonState] shouldEnable=${shouldEnable}, hasComponent=${state.hasComponent}, sizeConfirmed=${state.sizeConfirmed}, hasGridData=${hasGridData}, buttonText=${elements.createTableBtn.textContent}`);
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
      // Remove loading state from AI button
      const aiBtn = document.getElementById('generateTableFromPromptBtn') as HTMLButtonElement | null;
      if (aiBtn) {
        aiBtn.classList.remove('loading');
        aiBtn.disabled = false;
      }
      
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
      const bodyVisibilityKey = bodyAvailableProps.find(p => bodyPropertyTypes[p] === 'BOOLEAN' && (p.toLowerCase().includes('show') || p.toLowerCase().includes('text')) && !p.toLowerCase().includes('slot')) || '';
      const bodySlotKey = bodyAvailableProps.find(p => bodyPropertyTypes[p] === 'BOOLEAN' && p.toLowerCase().includes('slot')) || '';

      // Apply header cell properties
      for (let c = 1; c <= desiredCols; c++) {
        const key = `header-${c}`;
        const cellState = getCellState(key);
        // Clear ALL old text properties regardless of component type (to handle component variant changes)
        Object.keys(cellState.properties).forEach(prop => {
          if (prop.toLowerCase().includes('text') && !prop.toLowerCase().includes('slot') && !prop.toLowerCase().includes('second')) {
            delete cellState.properties[prop];
          }
        });
        cellState.properties[headerTextKey] = fixedHeaders[c - 1] || `H${c}`;
      }

      // Apply body cell properties
      for (let r = 0; r < desiredRows; r++) {
        for (let c = 0; c < desiredCols; c++) {
          const key = `${r + 1},${c + 1}`;
          const cellState = getCellState(key);
          // Clear ALL old text properties regardless of component type (to handle component variant changes)
          Object.keys(cellState.properties).forEach(prop => {
            if (prop.toLowerCase().includes('text') && !prop.toLowerCase().includes('slot') && !prop.toLowerCase().includes('second')) {
              delete cellState.properties[prop];
            }
          });
          // Clear any old boolean properties that will be set
          if (bodyVisibilityKey && typeof cellState.properties[bodyVisibilityKey] !== 'undefined') {
            delete cellState.properties[bodyVisibilityKey];
          }
          if (bodySlotKey && typeof cellState.properties[bodySlotKey] !== 'undefined') {
            delete cellState.properties[bodySlotKey];
          }
          cellState.properties[bodyTextKey] = fixedRows[r][c] || '';
          // Ensure text is visible in preview (and in final table) by enabling visibility
          if (bodyVisibilityKey) {
            cellState.properties[bodyVisibilityKey] = true;
          }
          // If a slot toggle exists, disable it to show text
          if (bodySlotKey) {
            cellState.properties[bodySlotKey] = false;
          }
        }
      }

      updateCellVisuals();
      state.sizeConfirmed = true; // Mark table as ready to create
      if (elements.createTableBtn) elements.createTableBtn.disabled = false;
      updateCreateButtonState(); // Update button state based on new data
      markChangesForReset(); // Enable reset button after AI content is applied
      showMessage(`AI content applied (clamped to ${desiredRows} rows x ${desiredCols} cols). Review/edit then click Create Table.`, 'success');
      
      // Uncheck the "Use single prompt" checkbox but preserve the prompt text
      // This prevents accidental re-generation while allowing users to view/edit their original prompt
      const useSinglePrompt = document.getElementById('useSinglePrompt') as HTMLInputElement | null;
      const singlePromptContainer = document.getElementById('singlePromptContainer') as HTMLElement | null;
      if (useSinglePrompt && useSinglePrompt.checked) {
        useSinglePrompt.checked = false;
        // Hide the container but DON'T clear the prompt text
        if (singlePromptContainer) {
          singlePromptContainer.classList.remove('show');
          setTimeout(() => {
            singlePromptContainer.style.display = 'none';
          }, 250);
        }
        console.log('ℹ️ [UI] Unchecked "Use single prompt" after AI generation (prompt text preserved)');
      }
      
      // Auto-analyze smart slots after main Watson AI table generation
      console.log('🤖 [UI] Auto-analyzing smart slots after main Watson AI table generation...');
      setTimeout(() => {
        analyzeSmartSlots(true);  // Pass true for auto-apply
      }, 500);
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
      // Ignore selection-cleared if we're about to receive table data
      // This prevents the UI from clearing when opening the plugin with a table already selected
      console.log('[UI] Received selection-cleared, but delaying UI clear to check for incoming table data...');
      
      // Wait a bit to see if we receive a table-related message
      setTimeout(() => {
        // Only clear if we still don't have a table after waiting
        if (!state.tableFrameId && !state.hasComponent) {
          state.hasComponent = false;
          showMessage("Please select a table cell component.", "error");
          elements.gridContainer.style.display = "none";
          elements.actionButtons.style.display = "none";
          state.selectedComponent = null;
          elements.landingPage.style.display = 'block';
          hideMessage(); // Hide status messages on landing page
          console.log('[UI] UI cleared after selection-cleared');
        } else {
          console.log('[UI] Ignored selection-cleared because table data arrived');
        }
      }, 100); // Wait 100ms for table messages to arrive
      break;

    case "table-created":
      hideLoader();
      // Remove loading state from button
      elements.createTableBtn.classList.remove('loading');
      elements.createTableBtn.disabled = false;
      
      if (msg.isComponent) {
        showMessage("Table component created successfully! You can now reuse it.", "success");
      } else {
        showMessage("Table created successfully!", "success");
      }
      // Show landing page and componentModeBtn again for new table generation
      elements.landingPage.style.display = 'flex';
      hideMessage(); // Hide status messages on landing page
      if (elements.componentModeBtn) {
        elements.componentModeBtn.style.display = 'inline-block';
      }
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
          // Prefer the first TEXT property; fallback to generic 'Cell text'
          return availableProps.find(p => propertyTypes[p] === 'TEXT') || 'Cell text';
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

        // Helper function to clean URLs by removing protocol-relative prefix
        const cleanUrl = (url: string): string => {
          if (url.startsWith('//')) {
            return url.substring(2); // Remove "//" prefix
          }
          return url;
        };

        const applyAiData = (cellKey: string, data: string) => {
          const cellState = getCellState(cellKey);
          const textProp = getCellTextProp();
          const visibilityProp = getTextVisibilityProp();

          console.log(`[applyAiData] Cell ${cellKey} BEFORE: properties =`, cellState.properties);
          
          if (textProp) {
            // Clean URL data before applying
            const cleanedData = cleanUrl(data);
            console.log(`🧹 Cleaned URL: "${data}" → "${cleanedData}"`);
            
            // Always apply AI data, clearing any existing text
            // This ensures the latest AI-generated data takes precedence
            cellState.properties[textProp] = cleanedData;
            console.log(`🔄 Applied AI data to cell ${cellKey}: "${cleanedData}" to property "${textProp}"`);
            console.log(`[applyAiData] Cell ${cellKey} AFTER: properties[${textProp}] =`, cellState.properties[textProp]);
          } else {
            console.warn(`⚠️ No textProp found for cell ${cellKey}`);
          }
          if (visibilityProp) {
            // Ensure the text is visible
            cellState.properties[visibilityProp] = true;
          }
          state.selectedCells.add(cellKey);
          
          // ✅ FIX: Explicitly save the updated cell state back to state.cellProperties
          state.cellProperties.set(cellKey, cellState);
          console.log(`✅ Saved updated cell state for ${cellKey} to state.cellProperties`);
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
        
        // Hide loader after AI generation completes
        hideLoader();
        console.log('[fake-data-generated] Hiding loader after AI generation');
        
        updateCellVisuals();
        closePropertyEditor();
        document.querySelectorAll('.cell').forEach(cell => {
          const divCell = cell as HTMLDivElement;
          const key = `${divCell.dataset.row},${divCell.dataset.col}`;
          if (state.selectedCells.has(key)) {
            divCell.classList.add('selected');
          }
        });
        // Ensure button is enabled after data generation
        state.sizeConfirmed = true;
        updateCreateButtonState();
        markChangesForReset();
        
        // Auto-analyze smart slots after Faker data is generated
        console.log('🤖 [UI] Auto-analyzing smart slots after Faker data generation...');
        setTimeout(() => {
          analyzeSmartSlots(true);  // Pass true for auto-apply
        }, 500);
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
        function getCellTextProp() { return availableProps.find(p => propertyTypes[p] === 'TEXT') || 'Cell text'; }
        function getTextVisibilityProp() { return availableProps.find(p => propertyTypes[p] === 'BOOLEAN' && (p.toLowerCase().includes('show') || p.toLowerCase().includes('text')) && !p.toLowerCase().includes('slot')); }
        
        // Helper function to clean URLs by removing protocol-relative prefix
        const cleanUrl = (url: string): string => {
          if (url.startsWith('//')) {
            return url.substring(2); // Remove "//" prefix
          }
          return url;
        };
        
        const applyAiData = (cellKey: string, data: string) => {
          const cellState = getCellState(cellKey);
          const textProp = getCellTextProp();
          const visibilityProp = getTextVisibilityProp();
          
          console.log(`[applyAiData - Watson] Cell ${cellKey} BEFORE: properties =`, cellState.properties);
          
          if (textProp) {
            // Clean URL data before applying
            const cleanedData = cleanUrl(data);
            console.log(`🧹 [Watson] Cleaned URL: "${data}" → "${cleanedData}"`);
            
            cellState.properties[textProp] = cleanedData;
            console.log(`🔄 Applied Watson AI data to cell ${cellKey}: "${cleanedData}" to property "${textProp}"`);
            console.log(`[applyAiData - Watson] Cell ${cellKey} AFTER: properties[${textProp}] =`, cellState.properties[textProp]);
          } else {
            console.warn(`⚠️ No textProp found for cell ${cellKey}`);
          }
          if (visibilityProp) cellState.properties[visibilityProp] = true;
          state.selectedCells.add(cellKey);
          
          // ✅ FIX: Explicitly save the updated cell state back to state.cellProperties
          state.cellProperties.set(cellKey, cellState);
          console.log(`✅ Saved updated cell state for ${cellKey} to state.cellProperties`);
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
        // Ensure button is enabled after data generation
        state.sizeConfirmed = true;
        updateCreateButtonState();
        markChangesForReset();
        
        // Auto-analyze smart slots after Watson data is generated
        console.log('🤖 [UI] Auto-analyzing smart slots after Watson data generation...');
        setTimeout(() => {
          analyzeSmartSlots(true);  // Pass true for auto-apply
        }, 500);
      }
      break;

    case "prompt-fallback-notice":
      hideLoader();
      showMessage(`Could not find a specific category for "${msg.prompt}". Using general text.`, "error");
      break;

    case "creation-error":
      hideLoader();
      // Remove loading state from button
      elements.createTableBtn.classList.remove('loading');
      elements.createTableBtn.disabled = false;
      
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

      // Only reset state if this is a NEW table scan (not a generated table being re-scanned)
      // Check if backend sent tableFrameId, indicating this is a generated table
      const isGeneratedTable = !!msg.details?.tableFrameId;
      if (!isGeneratedTable) {
        state.cellProperties.clear();
        state.selectedCells.clear();
        state.currentEditingCell = null;
        state.sizeConfirmed = false;
        state.tableFrameId = undefined;
        console.log('[DEBUG] Reset state for new Carbon table scan');
      } else {
        // Set tableFrameId for generated tables to prevent UI clear
        state.tableFrameId = msg.details.tableFrameId;
        console.log('[DEBUG] Skipped state reset - generated table detected (tableFrameId:', state.tableFrameId, ')');
      }

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
      
      // Initialize slot variables for all cells (same logic as in renderSlotComponentControls)
      initializeSlotVariables();
      // Automatically select all cells
      document.querySelectorAll('.cell').forEach(cell => {
        const key = `${(cell as HTMLDivElement).dataset.row},${(cell as HTMLDivElement).dataset.col}`;
        state.selectedCells.add(key);
        cell.classList.add('selected');
      });
      state.sizeConfirmed = true;
      // Show grid UI with animation
      elements.gridContainer.style.display = 'flex';
      elements.actionButtons.style.display = 'flex';
      requestAnimationFrame(() => {
        elements.gridContainer.classList.add('show');
      });

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
                <input type="checkbox" id="scanToolbarToggle" ${details.toolbar ? 'checked' : ''}>
                <label for="scanToolbarToggle">Toolbar</label>
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
      console.log(`[UI] Generated table selected but no settings available - metadata missing`);
      
      // Show error message and prompt to scan Carbon table first
      elements.landingPage.style.display = 'flex';
      elements.gridContainer.style.display = 'none';
      elements.actionButtons.style.display = 'none';
      
      showMessage(
        "⚠️ Table metadata not found. Please scan a Carbon Data Table component first to load the table structure, then select your generated table again.",
        "error"
      );
      
      state.hasComponent = false;
      state.tableFrameId = undefined;
      
      // Initialize slot variables for all cells
      initializeSlotVariables();

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

      // Show UI with animation
      elements.landingPage.style.display = 'none';
      elements.gridContainer.style.display = 'flex';
      elements.actionButtons.style.display = 'flex';
      requestAnimationFrame(() => {
        elements.gridContainer.classList.add('show');
      });

      console.log(`[UI] Showing grid and action buttons. Grid display: ${elements.gridContainer.style.display}, Action buttons display: ${elements.actionButtons.style.display}`);

      // Update state from settings
      state.gridRows = settings.rows || 5;
      state.gridCols = settings.columns || 5;
      
      // ✅ Update input fields immediately with correct dimensions
      console.log(`[UI] Updating input fields from edit-existing-table: ${settings.rows} rows × ${settings.columns} columns`);
      const rowsInput = document.getElementById('scanRowsInput') as HTMLInputElement;
      const colsInput = document.getElementById('scanColsInput') as HTMLInputElement;
      if (rowsInput) rowsInput.value = String(settings.rows || 5);
      if (colsInput) colsInput.value = String(settings.columns || 5);

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
                <div class="property-field checkbox">
                  <input type="checkbox" id="scanToolbarToggle" ${settings.includeToolbar ? 'checked' : ''}>
                  <label for="scanToolbarToggle">Toolbar</label>
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
      // Remove loading state from button
      elements.createTableBtn.classList.remove('loading');
      elements.createTableBtn.disabled = false;
      
      showMessage("Table updated successfully!", "success");
      // Keep the button as "Update Table" since we're still editing the same table
      elements.createTableBtn.textContent = 'Update Table';
      break;

    case "components-discovered":
      hideMessage();
      console.log('📦 Components discovered:', msg.components);
      
      if (msg.error) {
        showMessage(`Error discovering components: ${msg.error}`, 'error');
      } else {
        const componentCount = msg.components ? msg.components.length : 0;
        const pageComponentCount = msg.pageComponents || 0;
        const libraryComponentCount = msg.libraryComponents || 0;
        
        showMessage(`✅ Discovered ${componentCount} components (${libraryComponentCount} from libraries, ${pageComponentCount} on page)`, 'success');
        
        // Log some examples to console
        if (msg.components && msg.components.length > 0) {
          console.log('📋 Sample components:');
          msg.components.slice(0, 10).forEach((comp: any, index: number) => {
            console.log(`  ${index + 1}. ${comp.name} (${comp.width}x${comp.height})`);
          });
        }
      }
      break;

    case "component-key-found":
      hideMessage();
      console.log('🔑 Component key found:', msg.component);
      showMessage(`Component key: ${msg.component.key}`, 'success');
      break;

    case "carbon-keys-collected":
      hideMessage();
      console.log('🔑 Carbon keys collected:', msg.keys);
      if (msg.newKey) {
        showMessage(`✅ Collected: ${msg.newKey.name} (${msg.keys.length} total)`, 'success');
      } else {
        showMessage(`ℹ️ Already collected: ${msg.newKey?.name || 'component'} (${msg.keys.length} total)`, 'success');
      }
      break;

    case "component-info-unavailable":
      console.log(`[UI] ⚠️ Component info unavailable:`, msg.message);
      console.log('[UI] ℹ️ This error message will remain visible until a valid table is scanned (error messages do not auto-hide)');
      
      // Show error message
      elements.landingPage.style.display = 'flex';
      elements.gridContainer.style.display = 'none';
      elements.actionButtons.style.display = 'none';
      
      showMessage(
        msg.message || "⚠️ Component information unavailable. Please scan a Carbon Data Table component first to load the table structure.",
        "error"
      );
      
      state.hasComponent = false;
      state.tableFrameId = undefined;
      break;

    case "component-info":
      console.log(`[UI] Received component info:`, msg.component);
      if (msg.component) {
        state.selectedComponent = msg.component;
        
        // If actual dimensions are provided (from generated table metadata), use them
        // BUT only if we're not currently editing a table (to prevent overwriting new columns)
        if (msg.actualRows !== undefined && msg.actualCols !== undefined && !state.currentEditingCell) {
          console.log(`[UI] Updating grid dimensions from generated table metadata: ${msg.actualRows} rows × ${msg.actualCols} columns`);
          state.gridRows = msg.actualRows;
          state.gridCols = msg.actualCols;
          
          // Update input fields
          const rowsInput = document.getElementById('scanRowsInput') as HTMLInputElement;
          const colsInput = document.getElementById('scanColsInput') as HTMLInputElement;
          if (rowsInput) rowsInput.value = String(msg.actualRows);
          if (colsInput) colsInput.value = String(msg.actualCols);
          
          // Recreate grid with correct dimensions
          createGrid();
        }
        
        // Update visuals now that we have component info
        updateCellVisuals();
        
        // If we have saved cell properties from a generated table, restore them
        if (msg.savedCellProperties) {
          console.log(`[UI] Restoring ${Object.keys(msg.savedCellProperties).length} saved cell properties`);
          
          // Calculate actual dimensions from savedCellProperties if not provided by backend
          if (msg.actualRows === undefined || msg.actualCols === undefined) {
            let maxRow = 0;
            let maxCol = 0;
            
            for (const key of Object.keys(msg.savedCellProperties)) {
              // Parse body cell keys in format "row-col" (0-indexed in backend)
              if (key.includes('-') && !key.startsWith('header-')) {
                const [row, col] = key.split('-').map(Number);
                maxRow = Math.max(maxRow, row + 1);
                maxCol = Math.max(maxCol, col + 1);
              }
              // Parse header keys in format "header-N"
              if (key.startsWith('header-')) {
                const col = parseInt(key.split('-')[1]);
                maxCol = Math.max(maxCol, col);
              }
            }
            
            if (maxRow > 0 && maxCol > 0) {
              console.log(`[UI] Calculated dimensions from savedCellProperties: ${maxRow} rows × ${maxCol} columns`);
              state.gridRows = maxRow;
              state.gridCols = maxCol;
              
              // Update input fields
              const rowsInput = document.getElementById('scanRowsInput') as HTMLInputElement;
              const colsInput = document.getElementById('scanColsInput') as HTMLInputElement;
              if (rowsInput) rowsInput.value = String(maxRow);
              if (colsInput) colsInput.value = String(maxCol);
              
              // Recreate grid with correct dimensions
              createGrid();
            }
          }
          
          // Restore cell properties including slotComponentProps, statusIconText, statusIconType
          for (const [key, value] of Object.entries(msg.savedCellProperties)) {
            const cellData = value as any;
            const cellState = getCellState(key);
            
            // Restore basic properties
            if (cellData.properties) {
              cellState.properties = { ...cellData.properties };
            }
            
            // Restore Status Icon specific properties
            if (cellData.slotComponentProps) {
              cellState.slotComponentProps = cellData.slotComponentProps;
              
              // Infer and set slotComponentName from slotComponentProps for tooltip display
              if (!cellState.slotComponentName) {
                if (cellData.slotComponentProps.Status) {
                  cellState.slotComponentName = 'Status Icon';
                } else if (cellData.slotComponentProps.userName) {
                  cellState.slotComponentName = 'Slot Group (Avatar + Text / Edit + Delete)';
                } else if (cellData.slotComponentProps.editKey && cellData.slotComponentProps.deleteKey) {
                  cellState.slotComponentName = 'Slot Group (Avatar + Text / Edit + Delete)';
                } else if (cellData.slotComponentProps.tagText) {
                  cellState.slotComponentName = 'Tag Set';
                } else if (cellData.slotComponentProps.overflowActions) {
                  cellState.slotComponentName = 'Overflow Menu';
                } else if (cellData.slotComponentProps.editAction) {
                  cellState.slotComponentName = 'Edit Icon';
                } else if (cellData.slotComponentProps.deleteAction) {
                  cellState.slotComponentName = 'Delete Icon';
                }
              }
            }
            if (cellData.statusIconText) {
              cellState.statusIconText = cellData.statusIconText;
              if (!cellState.slotComponentName) {
                cellState.slotComponentName = 'Status Icon';
              }
            }
            if (cellData.statusIconType) {
              cellState.statusIconType = cellData.statusIconType;
            }
            // Restore slotComponentName if it was saved (this ensures persistence across saves)
            if (cellData.slotComponentName) {
              cellState.slotComponentName = cellData.slotComponentName;
            }
            if (cellData.customCellTextEnabled !== undefined) {
              cellState.customCellTextEnabled = cellData.customCellTextEnabled;
            }
            if (cellData.colWidth !== undefined) {
              cellState.colWidth = cellData.colWidth;
            }
          }
          
          console.log(`[UI] Finished restoring all cell properties`);
        } else {
          console.log(`[UI] No savedCellProperties received from backend`);
        }
        
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

    case 'smart-slot-suggestions':
      console.log('💡 [UI] Received smart slot suggestions:', msg.suggestions);
      if (msg.suggestions && msg.suggestions.length > 0) {
        console.log('📊 [UI] Smart Slot Suggestions:');
        msg.suggestions.forEach((suggestion: any, index: number) => {
          console.log(`  ${index + 1}. Column "${suggestion.columnName}" (${suggestion.columnIndex})`);
          console.log(`     Type: ${suggestion.contentType} → Suggested Component: ${suggestion.suggestedComponent}`);
          console.log(`     Samples: ${suggestion.samples.join(', ')}`);
          console.log(`     Confidence: ${(suggestion.confidence * 100).toFixed(0)}%`);
        });
        showMessage(`💡 Found ${msg.suggestions.length} smart slot suggestions! Check console.`, 'success');
      } else {
        console.log('ℹ️ [UI] No smart slot suggestions found');
        showMessage('No smart slot suggestions found for this data', 'success');
      }
      break;
    
    case 'auto-apply-smart-slots':
      console.log('🤖 [UI] Auto-applying smart slot suggestions:', msg.suggestions);
      
      // ✅ FIX: First, clear ALL existing slots from all cells to start fresh
      console.log('🧹 [UI] Clearing all existing slots before applying new ones...');
      for (let row = 1; row <= state.gridRows; row++) {
        for (let col = 1; col <= state.gridCols; col++) {
          const cellKey = `${row},${col}`;
          const cellState = getCellState(cellKey);
          
          // Clear slot properties
          if (state.selectedComponent?.availableProperties) {
            const slotProp = state.selectedComponent.availableProperties.find((p: string) => 
              p.toLowerCase().includes('slot') && !p.toLowerCase().includes('swap')
            );
            const swapSlotProp = state.selectedComponent.availableProperties.find((p: string) => 
              p.toLowerCase().includes('swap') && p.toLowerCase().includes('slot')
            );
            
            if (slotProp) {
              cellState.properties[slotProp] = false;
              cellState.slot = false;
            }
            if (swapSlotProp) {
              cellState.properties[swapSlotProp] = null;
            }
          }
          
          // Clear slot component properties
          cellState.slotComponentProps = null;
          cellState.slotComponentName = null;
          cellState.statusIconText = null;
          
          // Re-enable "Show text" since slot is disabled
          if (state.selectedComponent?.availableProperties) {
            const showTextProp = state.selectedComponent.availableProperties.find((p: string) => 
              p.toLowerCase().includes('show') && p.toLowerCase().includes('text')
            );
            if (showTextProp) {
              cellState.properties[showTextProp] = true;
              cellState.customCellTextEnabled = true;
            }
          }
          
          // Save the cleared state
          state.cellProperties.set(cellKey, cellState);
        }
      }
      console.log('✅ [UI] Cleared all existing slots');
      
      // Update grid visuals to reflect cleared slots
      updateCellVisuals();
      
      if (msg.suggestions && msg.suggestions.length > 0) {
        // Apply each suggestion to the corresponding column
        msg.suggestions.forEach((suggestion: any) => {
          const colIndex = suggestion.columnIndex;  // 0-based from backend
          const colNumber = colIndex + 1;  // Convert to 1-based for UI
          
          console.log(`🎯 [UI] Applying ${suggestion.suggestedComponent} to column ${colNumber} (${suggestion.columnName})`);
          
          if (suggestion.componentId) {
            console.log(`  📦 Component imported: ${suggestion.componentName} (ID: ${suggestion.componentId})`);
          }
          
          // Enable slot and set swap slot component for all cells in this column
          for (let row = 1; row <= state.gridRows; row++) {
            const cellKey = `${row},${colNumber}`;
            const cellState = getCellState(cellKey);
            
            // Get the current cell text to pass to the slotted component
            let cellText = '';
            if (cellState.properties) {
              const cellTextProp = Object.keys(cellState.properties).find(k => {
                const propValue = cellState.properties[k];
                return typeof propValue === 'string' && 
                       k.toLowerCase().includes('text') && 
                       !k.toLowerCase().includes('second') &&
                       !k.toLowerCase().includes('show');
              });
              if (cellTextProp) {
                cellText = String(cellState.properties[cellTextProp]);
              }
            }
            
            // Find the Slot and Swap slot properties
            if (state.selectedComponent?.availableProperties) {
              const slotProp = state.selectedComponent.availableProperties.find((p: string) => 
                p.toLowerCase().includes('slot') && !p.toLowerCase().includes('swap')
              );
              const swapSlotProp = state.selectedComponent.availableProperties.find((p: string) => 
                p.toLowerCase().includes('swap') && p.toLowerCase().includes('slot')
              );
              
              if (slotProp) {
                cellState.properties[slotProp] = true;
                // Also set the dedicated slot variable
                cellState.slot = true;
                console.log(`  ✅ Enabled slot for cell ${cellKey} (set both property and slot variable)`);
              }
              
              // Disable "Show text" for all slots (including status icons)
              // The slot component itself will handle displaying text
              const showTextProp = state.selectedComponent.availableProperties.find((p: string) => 
                p.toLowerCase().includes('show') && p.toLowerCase().includes('text')
              );
              if (showTextProp) {
                cellState.properties[showTextProp] = false;
                cellState.customCellTextEnabled = false;
                console.log(`  🔇 Disabled "Show text" for cell ${cellKey}`);
              }
              
              // If we have a component ID, set the Swap slot property
              if (swapSlotProp && suggestion.componentId) {
                cellState.properties[swapSlotProp] = suggestion.componentId;
                
                // Hardcode component names based on suggestion type for tooltip display
                const componentNameMap: Record<string, string> = {
                  'statusIcon': 'Status Icon',
                  'slotGroup': 'Slot Group (Avatar + Text / Edit + Delete)',
                  'tag': 'Tag Set',
                  'overflow': 'Overflow Menu',
                  'edit': 'Edit Icon',
                  'delete': 'Delete Icon',
                  'checkbox': 'Checkbox',
                  'link': 'Link',
                  'avatar': 'Avatar'
                };
                
                cellState.slotComponentName = componentNameMap[suggestion.suggestedComponent] || suggestion.componentName || 'Slot Component';
                console.log(`  🔄 Set swap slot to ${cellState.slotComponentName} for cell ${cellKey}`);
                
                // For Status Icon: map cell text to the appropriate status type text property
                if (suggestion.suggestedComponent === 'statusIcon' && cellText) {
                  // Map common status values to Carbon status types
                  const statusMap: Record<string, string> = {
                    'failed': 'Failed',
                    'succeeded': 'Succeeded',
                    'success': 'Succeeded',
                    'pending': 'In-progress',
                    'pending verification': 'In-progress',
                    'pending approval': 'In-progress',
                    'pending review': 'In-progress',
                    'in-progress': 'In-progress',
                    'in progress': 'In-progress',
                    'processing': 'In-progress',
                    'verifying': 'In-progress',
                    'under review': 'In-progress',
                    'active': 'Succeeded',        // Active should be green/success
                    'inactive': 'Caution major',  // Inactive should be warning
                    'premium': 'Succeeded',       // Premium should be green/success
                    'normal': 'Normal',
                    'completed': 'Succeeded',
                    'incomplete': 'Caution major',
                    'not started': 'Unknown',
                    'unknown': 'Unknown',
                    'warning': 'Caution minor',
                    'probation': 'Caution major',
                    'undefined': 'Undefined',
                    'approved': 'Succeeded',
                    'rejected': 'Failed',
                    'accepted': 'Succeeded',      // ✅ Added: Accepted = Success
                    'declined': 'Failed',         // ✅ Added: Declined = Failed
                    'cancelled': 'Failed',
                    'draft': 'Caution minor',
                    'published': 'Succeeded',
                    'archived': 'Caution minor',
                    'enabled': 'Succeeded',
                    'disabled': 'Caution major',
                    'open': 'In-progress',        // ✅ Added: Open tickets/issues
                    'closed': 'Succeeded',        // ✅ Added: Closed = Complete
                    'confirmed': 'Succeeded',     // ✅ Added: Confirmed = Success
                    'denied': 'Failed',           // ✅ Added: Denied = Failed
                    'resolved': 'Succeeded',      // ✅ Added: Resolved = Success
                    'unresolved': 'Caution major' // ✅ Added: Unresolved = Warning
                  };
                  
                  // Find the matching status type (case-insensitive)
                  const cellTextLower = cellText.toLowerCase();
                  const statusType = statusMap[cellTextLower] || 'Normal'; // Default to Normal
                  
                  console.log(`🔍 [STATUS MAPPING DEBUG] Cell text: "${cellText}" → Lowercase: "${cellTextLower}" → Mapped to: "${statusType}"`);
                  console.log(`🔍 [STATUS MAPPING DEBUG] Available mappings:`, Object.keys(statusMap));
                  
                  // Store the properties for the slotted Status Icon
                  // We need to set BOTH the Status property AND the text property for that status
                  cellState.slotComponentProps = {
                    'Status': statusType,           // Set the status type (Failed, Normal, In-progress, etc.)
                    'Label': true,                   // Enable label toggle
                    [`${statusType} text`]: cellText // Set the text for the specific status type
                  };
                  
                  // Store the original text and type for editing
                  cellState.statusIconText = cellText;
                  cellState.statusIconType = statusType;
                  
                  console.log(`  📝 Set Status Icon: status="${statusType}", text="${cellText}"`);
                  console.log(`  📦 [UI] slotComponentProps stored:`, cellState.slotComponentProps);
                }
                
                // For Edit/Delete actions (2 actions): store slot group information
                // CHECK THIS FIRST before general slotGroup to avoid mixing with user names
                else if (suggestion.contentType === 'editDelete') {
                  // Store the information needed for nested slot swapping with Edit + Delete icons
                  cellState.slotComponentProps = {
                    'nestedSlots': true,  // Flag to indicate this needs nested slot handling
                    'editKey': 'a4ba4c4aa1f2b0f0a5206341aafbb7d7eafa47e6',  // Edit component key
                    'deleteKey': '84a7c6755b83b8e88ca803851c280d1e06255b93'  // Delete component key
                  };
                  
                  console.log(`  ⚡ Set Edit/Delete actions (2 actions) for cell ${cellKey}`);
                  console.log(`  📦 [UI] slotComponentProps stored:`, cellState.slotComponentProps);
                }
                
                // For Slot Group (Avatar + Text for user names): store nested slot information
                else if (suggestion.suggestedComponent === 'slotGroup' && cellText) {
                  // Store the information needed for nested slot swapping
                  // The backend will handle swapping the nested slots
                  cellState.slotComponentProps = {
                    'nestedSlots': true,  // Flag to indicate this needs nested slot handling
                    'userName': cellText,  // The name to display
                    'avatarKey': 'd80f0d175851756c4601e87e6e0abeb5539b24e8',  // Avatar component key
                    'textKey': 'e73c62eb16dcb7f54df1384f176fc7a3c0f64df5'     // Text component key
                  };
                  
                  console.log(`  👤 Set Slot Group: userName="${cellText}"`);
                  console.log(`  📦 [UI] slotComponentProps stored:`, cellState.slotComponentProps);
                }
                
                // For Tag: store tag information for configuration
                else if (suggestion.suggestedComponent === 'tag' && cellText) {
                  // Store the information needed for tag configuration
                  // The backend will handle configuring the tag text and colors
                  cellState.slotComponentProps = {
                    'suggestedComponent': 'tag',  // Flag to indicate this is a tag component
                    'tagText': cellText  // The text to display in tags
                  };
                  
                  console.log(`  🏷️ Set Tag: tagText="${cellText}"`);
                  console.log(`  📦 [UI] slotComponentProps stored:`, cellState.slotComponentProps);
                }
                
                // For Overflow Menu: store configuration flag for backend
                else if (suggestion.suggestedComponent === 'overflow') {
                  // Store the information needed for overflow menu configuration
                  // The backend will handle setting the Content group alignment to top right
                  cellState.slotComponentProps = {
                    'suggestedComponent': 'overflow'  // Flag to indicate this is an overflow menu
                  };
                  console.log(`  📋 Set Overflow Menu for cell ${cellKey}`);
                  console.log(`  📦 [UI] slotComponentProps stored:`, cellState.slotComponentProps);
                }
                
                // For single Edit action: simple icon swap (no slot group needed)
                else if (suggestion.suggestedComponent === 'edit') {
                  // No special configuration needed - just the swap
                  console.log(`  ✏️ Set single Edit icon for cell ${cellKey}`);
                }
                
                // For single Delete action: simple icon swap (no slot group needed)
                else if (suggestion.suggestedComponent === 'delete') {
                  // No special configuration needed - just the swap
                  console.log(`  🗑️ Set single Delete icon for cell ${cellKey}`);
                }
                
                // For Link: store link information for configuration
                else if (suggestion.suggestedComponent === 'link' && cellText) {
                  // Store the information needed for link configuration
                  // The backend will handle configuring the link text and properties
                  cellState.slotComponentProps = {
                    'suggestedComponent': 'link',  // Flag to indicate this is a link component
                    'linkText': cellText  // The text to display in the link
                  };
                  console.log(`  🔗 Set Link: linkText="${cellText}"`);
                  console.log(`  📦 [UI] slotComponentProps stored:`, cellState.slotComponentProps);
                }
              }
            }
          }
          
          // Update cell visuals to show the changes
          updateCellVisuals();
        });
        
        console.log(`✨ [UI] Auto-applied ${msg.suggestions.length} smart slot suggestion(s)`);
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
  
  // Auto-analyze smart slots after data is loaded
  console.log('🤖 [UI] Auto-analyzing smart slots after data load...');
  setTimeout(() => {
    analyzeSmartSlots(true);  // Pass true for auto-apply
  }, 500);  // Small delay to ensure grid is fully rendered
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
      const columnData: { rowIndex: number; value: string }[] = [];
      
      for (let r = 1; r <= rows; r++) {
        const cellData = sortedCellProperties.get(`${r},${c}`);
        let cellValue = '';
        if (cellData?.properties) {
          if (typeof cellData.properties['Cell text'] === 'string' && cellData.properties['Cell text'].toString().trim() !== '') {
            cellValue = cellData.properties['Cell text'];
          } else {
            const textProp = Object.keys(cellData.properties).find(prop => 
              prop.toLowerCase().includes('text') && !prop.toLowerCase().includes('second') && typeof cellData.properties[prop] === 'string');
            cellValue = (textProp && cellData.properties[textProp]) || cellData.properties['Cell text#12234:32'] || '';
          }
        }
        columnData.push({ rowIndex: r, value: String(cellValue) });
      }
      
      columnData.sort((a, b) => {
        const numA = parseFloat(a.value);
        const numB = parseFloat(b.value);
        if (!isNaN(numA) && !isNaN(numB)) {
          return sorted === 'Ascending' ? numA - numB : numB - numA;
        }
        return sorted === 'Ascending' ? a.value.localeCompare(b.value) : b.value.localeCompare(a.value);
      });
      
      // Reorder entire rows based on the sorted order of the selected column
      const newMap = new Map(sortedCellProperties);
      for (let newRow = 1; newRow <= rows; newRow++) {
        const sourceRow = columnData[newRow - 1].rowIndex;
        for (let cc = 1; cc <= cols; cc++) {
          const src = sortedCellProperties.get(`${sourceRow},${cc}`);
          newMap.set(`${newRow},${cc}`, src);
        }
      }
      // Replace with reordered map and stop after first sortable column
      return newMap;
    }
  }
  
  return sortedCellProperties;
}

// Column Width Management Functions
function calculateTotalWidth(): number {
  const columnList = document.getElementById('columnList');
  if (!columnList) return 0;
  
  const widthInputs = columnList.querySelectorAll('.column-width-input:not(.deleted .column-width-input)') as NodeListOf<HTMLInputElement>;
  let total = 0;
  
  widthInputs.forEach(input => {
    const columnIndex = parseInt(input.dataset.column || '1');
    const width = parseInt(input.value) || getDefaultColumnWidthUI(columnIndex);
    total += width;
  });
  
  console.log('[calculateTotalWidth] Total width:', total);
  return total;
}

function getColumnProportions(): number[] {
  const columnList = document.getElementById('columnList');
  if (!columnList) return [];
  
  const widthInputs = columnList.querySelectorAll('.column-width-input:not(.deleted .column-width-input)') as NodeListOf<HTMLInputElement>;
  const widths: number[] = [];
  let total = 0;
  
  widthInputs.forEach(input => {
    const columnIndex = parseInt(input.dataset.column || '1');
    const width = parseInt(input.value) || getDefaultColumnWidthUI(columnIndex);
    widths.push(width);
    total += width;
  });
  
  // Calculate proportions (as percentages)
  const proportions = widths.map(w => total > 0 ? w / total : 1 / widths.length);
  console.log('[getColumnProportions] Proportions:', proportions);
  return proportions;
}

function distributeWidthProportionally(totalWidth: number) {
  const columnList = document.getElementById('columnList');
  if (!columnList || totalWidth <= 0) return;
  
  const proportions = getColumnProportions();
  const widthInputs = columnList.querySelectorAll('.column-width-input:not(.deleted .column-width-input)') as NodeListOf<HTMLInputElement>;
  
  console.log('[distributeWidthProportionally] Distributing total width:', totalWidth);
  
  widthInputs.forEach((input, index) => {
    if (index < proportions.length) {
      const newWidth = Math.round(totalWidth * proportions[index]);
      input.value = String(Math.max(50, newWidth)); // Minimum 50px
      console.log(`  Column ${index + 1}: ${newWidth}px (${(proportions[index] * 100).toFixed(1)}%)`);
    }
  });
}

function handleColumnWidthChange(event: Event) {
  const input = event.target as HTMLInputElement;
  const value = input.value;
  
  // Allow empty string during editing
  if (value === '') {
    return;
  }
  
  const newWidth = parseInt(value);
  
  // Only validate if it's a valid number
  if (!isNaN(newWidth)) {
    // Update total width (even if below minimum, to show live calculation)
    const totalWidthInput = document.getElementById('totalTableWidthInput') as HTMLInputElement;
    if (totalWidthInput) {
      const newTotal = calculateTotalWidth();
      totalWidthInput.value = String(newTotal);
      console.log('[handleColumnWidthChange] Updated total width to:', newTotal);
    }
  }
}

function handleColumnWidthBlur(event: Event) {
  const input = event.target as HTMLInputElement;
  const newWidth = parseInt(input.value);
  
  // Enforce minimum width on blur
  if (isNaN(newWidth) || newWidth < 50) {
    input.value = '50';
    // Update total width after correction
    const totalWidthInput = document.getElementById('totalTableWidthInput') as HTMLInputElement;
    if (totalWidthInput) {
      const newTotal = calculateTotalWidth();
      totalWidthInput.value = String(newTotal);
    }
  }
}

function handleTotalWidthChange(event: Event) {
  const input = event.target as HTMLInputElement;
  const value = input.value;
  
  // Allow empty string during editing
  if (value === '') {
    return;
  }
  
  const newTotal = parseInt(value);
  
  // Only validate and distribute if it's a valid number
  if (!isNaN(newTotal) && newTotal >= 100) {
    console.log('[handleTotalWidthChange] New total width:', newTotal);
    distributeWidthProportionally(newTotal);
  }
}

function handleTotalWidthBlur(event: Event) {
  const input = event.target as HTMLInputElement;
  const newTotal = parseInt(input.value);
  
  // Enforce minimum on blur
  if (isNaN(newTotal) || newTotal < 100) {
    input.value = '100';
    distributeWidthProportionally(100);
  }
}

function applyColumnWidths() {
  const columnList = document.getElementById('columnList');
  if (!columnList) return;
  
  const columnItems = Array.from(columnList.children) as HTMLElement[];
  
  console.log('[applyColumnWidths] Applying column widths to cell properties');
  
  columnItems.forEach(item => {
    if (item.classList.contains('deleted')) return;
    
    const columnIndex = parseInt(item.dataset.columnIndex || '0');
    const widthInput = item.querySelector('.column-width-input') as HTMLInputElement;
    const width = parseInt(widthInput?.value) || getDefaultColumnWidthUI(columnIndex);
    
    console.log(`  Column ${columnIndex}: ${width}px`);
    
    // Apply width to all cells in this column
    for (let row = 1; row <= state.gridRows; row++) {
      const cellKey = `${row},${columnIndex}`;
      const cellState = getCellState(cellKey);
      cellState.colWidth = width;
    }
    
    // Also apply to header cell
    const headerKey = `header-${columnIndex}`;
    const headerState = getCellState(headerKey);
    headerState.colWidth = width;
  });
  
  console.log('[applyColumnWidths] Column widths applied successfully');
}

// Column Reordering Functions
function openColumnReorderModal() {
  const overlay = document.getElementById('columnReorderOverlay');
  const columnList = document.getElementById('columnList');
  const totalWidthInput = document.getElementById('totalTableWidthInput') as HTMLInputElement;
  
  if (!overlay || !columnList) return;
  
  // Clear existing column items
  columnList.innerHTML = '';
  
  // Create column items based on current grid columns
  for (let c = 1; c <= state.gridCols; c++) {
    const columnItem = createColumnItem(c);
    columnList.appendChild(columnItem);
  }
  
  // Calculate and set initial total width
  const totalWidth = calculateTotalWidth();
  if (totalWidthInput) {
    totalWidthInput.value = String(totalWidth);
  }
  
  // Show modal with animation
  overlay.style.display = 'flex';
  requestAnimationFrame(() => {
    overlay.classList.add('show');
  });
  
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
  
  // Get column width from any cell in this column (check first row cells)
  let columnWidth = getDefaultColumnWidthUI(columnIndex); // Smart default width
  for (let row = 1; row <= state.gridRows; row++) {
    const cellKey = `${row},${columnIndex}`;
    const cellData = state.cellProperties.get(cellKey);
    if (cellData && cellData.colWidth) {
      columnWidth = cellData.colWidth;
      break;
    }
  }
  
  item.innerHTML = `
    <span class="drag-handle">⋮⋮</span>
    <span class="column-name">${columnName}</span>
    <span class="column-index">${columnIndex}</span>
    <div class="column-width-wrapper">
      <input type="number" class="column-width-input" data-column="${columnIndex}" value="${columnWidth}" step="10">
      <span class="width-unit">px</span>
    </div>
    <button class="column-delete-btn" title="Delete column">×</button>
  `;
  
  // Add delete functionality
  const deleteBtn = item.querySelector('.column-delete-btn') as HTMLButtonElement;
  deleteBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    item.classList.toggle('deleted');
    item.draggable = !item.classList.contains('deleted');
  });
  
  // Add width input event listeners
  const widthInput = item.querySelector('.column-width-input') as HTMLInputElement;
  widthInput.addEventListener('input', handleColumnWidthChange);
  widthInput.addEventListener('blur', handleColumnWidthBlur);
  widthInput.addEventListener('click', (e) => e.stopPropagation());
  
  // Add drag event listeners
  item.addEventListener('dragstart', handleDragStart);
  item.addEventListener('dragover', handleDragOver);
  item.addEventListener('drop', handleDrop);
  item.addEventListener('dragend', handleDragEnd);
  
  return item;
}

// Flag to prevent duplicate event listener setup
let modalListenersSetup = false;

function setupColumnReorderModalListeners() {
  // Prevent duplicate event listener setup
  if (modalListenersSetup) {
    console.log('🔧 [setupColumnReorderModalListeners] Event listeners already setup, skipping...');
    return;
  }
  
  const closeBtn = document.getElementById('closeReorderModal');
  const cancelBtn = document.getElementById('cancelReorderBtn');
  const applyBtn = document.getElementById('applyReorderBtn');
  const overlay = document.getElementById('columnReorderOverlay');
  const totalWidthInput = document.getElementById('totalTableWidthInput') as HTMLInputElement;
  
  // Close modal handlers with animation
  const closeModal = () => {
    if (overlay) {
      overlay.classList.remove('show');
      setTimeout(() => {
        overlay.style.display = 'none';
      }, 250);
    }
  };
  
  closeBtn?.addEventListener('click', closeModal);
  cancelBtn?.addEventListener('click', closeModal);
  
  // Apply reorder and width changes handler
  applyBtn?.addEventListener('click', () => {
    applyColumnWidths(); // Apply width changes first
    applyColumnReorder(); // Then apply reordering
  });
  
  // Total width input listeners
  if (totalWidthInput) {
    totalWidthInput.addEventListener('input', handleTotalWidthChange);
    totalWidthInput.addEventListener('blur', handleTotalWidthBlur);
  }
  
  // Close on overlay click
  overlay?.addEventListener('click', (e) => {
    if (e.target === overlay) closeModal();
  });
  
  // Mark as setup
  modalListenersSetup = true;
  console.log('🔧 [setupColumnReorderModalListeners] Event listeners setup complete');
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
  
  console.log('🔀 [applyColumnReorder] Column items in DOM order:', columnItems.map(item => ({
    index: item.dataset.columnIndex,
    name: item.querySelector('.column-name')?.textContent,
    deleted: item.classList.contains('deleted')
  })));
  console.log('🔀 [applyColumnReorder] Remaining columns (original indices):', remainingColumns);
  
  // Apply the reordering and deletion to the table data
  reorderAndDeleteColumns(remainingColumns);
  
  // Close modal with animation
  const overlay = document.getElementById('columnReorderOverlay');
  if (overlay) {
    overlay.classList.remove('show');
    setTimeout(() => {
      overlay.style.display = 'none';
    }, 250);
  }
  
  // Show success message
  const deletedCount = columnItems.length - remainingColumns.length;
  const message = deletedCount > 0 
    ? `Columns reordered and ${deletedCount} column(s) deleted successfully!`
    : 'Columns reordered successfully!';
  showMessage(message, 'success');
}

function reorderAndDeleteColumns(remainingColumns: number[]) {
  console.log('🔄 [reorderAndDeleteColumns] Starting reorder with columns:', remainingColumns);
  
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
      console.log(`🔄 [reorderAndDeleteColumns] Moved header ${oldCol} → ${newCol} (${oldKey} → ${newKey})`);
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
  
  // Update cell visuals to show the reordered content
  updateCellVisuals();
  
  // Mark changes for reset
  markChangesForReset();
}

// Make openColumnReorderModal available globally
(window as any).openColumnReorderModal = openColumnReorderModal;

export { }; // Treat this file as a module
