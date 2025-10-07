import type { State, Elements, ComponentInfo } from '../types';
import { showMessage, hideMessage, showLoader, hideLoader } from '../utils/helpers';
import { sortColumnData } from '../utils/sorting';

// Message type constants
export const MESSAGE_TYPES = {
  AI_TABLE_RESPONSE: 'ai-table-response',
  TABLE_SELECTED: 'table-selected',
  COMPONENT_SELECTED: 'component-selected',
  SELECTION_CLEARED: 'selection-cleared',
  TABLE_CREATED: 'table-created',
  FAKE_DATA_RESPONSE: 'fake-data-response',
  WATSONX_DATA_RESPONSE: 'watsonx-data-response',
  PROMPT_FALLBACK_NOTICE: 'prompt-fallback-notice',
  CREATION_ERROR: 'creation-error',
  COMPONENT_PROPERTIES: 'component-properties',
  SCAN_TABLE: 'scan-table',
  SCAN_TABLE_RESULT: 'scan-table-result',
  GENERATED_TABLE_SELECTED_NO_SETTINGS: 'generated-table-selected-no-settings',
  EDIT_EXISTING_TABLE: 'edit-existing-table',
  TABLE_UPDATED: 'table-updated',
  COMPONENT_INFO: 'component-info',
  WATSONX_SETTINGS: 'watsonx-settings',
  // Outgoing message types
  GENERATE_TABLE_WITH_AI: 'generate-table-with-ai',
  SAVE_WATSONX_API_KEY: 'save-watsonx-api-key',
  RESIZE: 'resize',
  LOAD_WATSONX_SETTINGS: 'load-watsonx-settings',
  REQUEST_COMPONENT_INFO: 'request-component-info',
  GENERATE_FAKE_DATA: 'generate-fake-data',
  GENERATE_WATSONX_DATA: 'generate-watsonx-data',
  REQUEST_SELECTION_STATE: 'request-selection-state'
} as const;

// Message interfaces
export interface PluginMessage {
  type: string;
  [key: string]: any;
}

export interface MessageEvent {
  data: {
    pluginMessage: PluginMessage;
  };
}

export type MessageHandlerFunction = (message: PluginMessage) => void;

/**
 * MessageHandler component for centralized plugin communication
 * Handles all incoming and outgoing messages between the UI and Figma plugin
 */
export class MessageHandler {
  private state: State;
  private elements: Elements;
  private messageHandlers: Map<string, MessageHandlerFunction> = new Map();
  
  // External dependencies that will be injected
  private gridManager: any = null;
  private propertyEditor: any = null;
  private pendingFakerContext: any = null;
  private updateCellVisuals: (() => void) | null = null;

  constructor(state: State, elements: Elements) {
    this.state = state;
    this.elements = elements;
    this.setupMessageListener();
  }

  /**
   * Set external dependencies
   */
  setDependencies(dependencies: {
    gridManager?: any;
    propertyEditor?: any;
    getCellState?: (key: string) => any;
    updateCellVisuals?: () => void;
    markChangesForReset?: () => void;
    clearChangesForReset?: () => void;
    setMode?: (mode: State['mode']) => void;
    updateCreateButtonState?: () => void;
    updateModeDependentVisibility?: () => void;
    renderHeaderFooterGrids?: () => void;
    closePropertyEditor?: () => void;
  }): void {
    Object.assign(this, dependencies);
  }

  /**
   * Set pending faker context (used by fake data handlers)
   */
  setPendingFakerContext(context: any): void {
    this.pendingFakerContext = context;
  }

  /**
   * Get pending faker context
   */
  getPendingFakerContext(): any {
    return this.pendingFakerContext;
  }

  /**
   * Clear pending faker context
   */
  clearPendingFakerContext(): void {
    this.pendingFakerContext = null;
  }

  /**
   * Set up the main window message listener
   */
  private setupMessageListener(): void {
    window.onmessage = (event: MessageEvent) => {
      this.handleIncomingMessage(event);
    };
  }

  /**
   * Handle incoming messages from the Figma plugin
   * Routes messages to appropriate handler methods
   */
  handleIncomingMessage(event: MessageEvent): void {
    // Check if DOM is ready and required elements exist
    const domReady = document.readyState === 'complete' || document.readyState === 'interactive';
    if (!domReady || !this.elements.grid) return;

    const msg = event.data.pluginMessage;
    if (!msg) return;

    console.log(`[UI] Received message: ${msg.type}`, msg);

    try {
      // Check if there's a registered handler for this message type
      const handler = this.messageHandlers.get(msg.type);
      if (handler) {
        handler(msg);
      } else {
        console.warn(`[MessageHandler] No handler registered for message type: ${msg.type}`);
      }
    } catch (error) {
      console.error(`[MessageHandler] Error handling message ${msg.type}:`, error);
    }
  }

  /**
   * Send a message to the Figma plugin
   */
  sendMessage(type: string, data?: any): void {
    const message = {
      pluginMessage: {
        type,
        ...data
      }
    };

    console.log(`[UI] Sending message: ${type}`, message);
    parent.postMessage(message, '*');
  }

  /**
   * Register a message handler for a specific message type
   */
  registerMessageHandler(type: string, handler: MessageHandlerFunction): void {
    this.messageHandlers.set(type, handler);
  }

  /**
   * Unregister a message handler
   */
  unregisterMessageHandler(type: string): void {
    this.messageHandlers.delete(type);
  }

  /**
   * Send scan table message
   */
  scanTable(): void {
    this.sendMessage(MESSAGE_TYPES.SCAN_TABLE);
  }

  /**
   * Send generate table with AI message
   */
  generateTableWithAI(prompt: string, rows: number, cols: number): void {
    this.sendMessage(MESSAGE_TYPES.GENERATE_TABLE_WITH_AI, {
      prompt,
      rows,
      cols
    });
  }

  /**
   * Save watsonx API key
   */
  saveWatsonxApiKey(apiKey: string): void {
    this.sendMessage(MESSAGE_TYPES.SAVE_WATSONX_API_KEY, { apiKey });
  }

  /**
   * Send resize message
   */
  resize(size: { w: number; h: number }): void {
    this.sendMessage(MESSAGE_TYPES.RESIZE, { size });
  }

  /**
   * Load watsonx settings
   */
  loadWatsonxSettings(): void {
    this.sendMessage(MESSAGE_TYPES.LOAD_WATSONX_SETTINGS);
  }

  /**
   * Request component info
   */
  requestComponentInfo(tableId?: string): void {
    this.sendMessage(MESSAGE_TYPES.REQUEST_COMPONENT_INFO, tableId ? { tableId } : undefined);
  }

  /**
   * Generate fake data
   */
  generateFakeData(dataType: string, count: number): void {
    this.sendMessage(MESSAGE_TYPES.GENERATE_FAKE_DATA, { dataType, count });
  }

  /**
   * Generate watsonx data
   */
  generateWatsonxData(options: {
    prompt: string;
    endpoint: string;
    apiKey: string;
    accessToken?: string;
    useAccessToken: boolean;
    useProxy: boolean;
    proxyUrl: string;
    count: number;
    remember: boolean;
  }): void {
    this.sendMessage(MESSAGE_TYPES.GENERATE_WATSONX_DATA, options);
  }

  /**
   * Request selection state
   */
  requestSelectionState(): void {
    this.sendMessage(MESSAGE_TYPES.REQUEST_SELECTION_STATE);
  }

  /**
   * Create table message
   */
  createTable(tableData: any): void {
    this.sendMessage('create-table', tableData);
  }

  /**
   * Initialize default message handlers
   * This method should be called after the MessageHandler is created
   */
  initializeDefaultHandlers(): void {
    // Register all the default message handlers
    this.registerMessageHandler(MESSAGE_TYPES.AI_TABLE_RESPONSE, this.handleAiTableResponse.bind(this));
    this.registerMessageHandler(MESSAGE_TYPES.TABLE_SELECTED, this.handleTableSelected.bind(this));
    this.registerMessageHandler(MESSAGE_TYPES.COMPONENT_SELECTED, this.handleComponentSelected.bind(this));
    this.registerMessageHandler(MESSAGE_TYPES.SELECTION_CLEARED, this.handleSelectionCleared.bind(this));
    this.registerMessageHandler(MESSAGE_TYPES.TABLE_CREATED, this.handleTableCreated.bind(this));
    this.registerMessageHandler(MESSAGE_TYPES.FAKE_DATA_RESPONSE, this.handleFakeDataResponse.bind(this));
    this.registerMessageHandler(MESSAGE_TYPES.WATSONX_DATA_RESPONSE, this.handleWatsonxDataResponse.bind(this));
    this.registerMessageHandler(MESSAGE_TYPES.PROMPT_FALLBACK_NOTICE, this.handlePromptFallbackNotice.bind(this));
    this.registerMessageHandler(MESSAGE_TYPES.CREATION_ERROR, this.handleCreationError.bind(this));
    this.registerMessageHandler(MESSAGE_TYPES.COMPONENT_PROPERTIES, this.handleComponentProperties.bind(this));
    this.registerMessageHandler(MESSAGE_TYPES.SCAN_TABLE, this.handleScanTable.bind(this));
    this.registerMessageHandler(MESSAGE_TYPES.SCAN_TABLE_RESULT, this.handleScanTableResult.bind(this));
    this.registerMessageHandler(MESSAGE_TYPES.GENERATED_TABLE_SELECTED_NO_SETTINGS, this.handleGeneratedTableSelectedNoSettings.bind(this));
    this.registerMessageHandler(MESSAGE_TYPES.EDIT_EXISTING_TABLE, this.handleEditExistingTable.bind(this));
    this.registerMessageHandler(MESSAGE_TYPES.TABLE_UPDATED, this.handleTableUpdated.bind(this));
    this.registerMessageHandler(MESSAGE_TYPES.COMPONENT_INFO, this.handleComponentInfo.bind(this));
    this.registerMessageHandler(MESSAGE_TYPES.WATSONX_SETTINGS, this.handleWatsonxSettings.bind(this));
  }

  // Message handler methods
  private handleAiTableResponse(msg: PluginMessage): void {
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
      return;
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
    const desiredRows = this.state.gridRows;
    const desiredCols = this.state.gridCols;

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
    this.gridManager?.createGrid();
    this.elements.gridContainer.style.display = 'flex';
    this.elements.actionButtons.style.display = 'flex';

    // Use header cell component properties if available, otherwise fallback to body cell properties
    const headerAvailableProps: string[] = this.state.headerCellComponent?.availableProperties || this.state.selectedComponent?.availableProperties || [];
    const headerPropertyTypes: { [key: string]: any } = this.state.headerCellComponent?.propertyTypes || this.state.selectedComponent?.propertyTypes || {};
    const bodyAvailableProps: string[] = this.state.selectedComponent?.availableProperties || [];
    const bodyPropertyTypes: { [key: string]: any } = this.state.selectedComponent?.propertyTypes || {};

    const headerTextKey = headerAvailableProps.find(p => headerPropertyTypes[p] === 'TEXT') || 'Cell text';
    const bodyTextKey = bodyAvailableProps.find(p => bodyPropertyTypes[p] === 'TEXT') || 'Cell text';
    const bodyVisibilityKey = bodyAvailableProps.find(p => bodyPropertyTypes[p] === 'BOOLEAN' && (p.toLowerCase().includes('show') || p.toLowerCase().includes('text')) && !p.toLowerCase().includes('slot')) || '';
    const bodySlotKey = bodyAvailableProps.find(p => bodyPropertyTypes[p] === 'BOOLEAN' && p.toLowerCase().includes('slot')) || '';

    // Apply header cell properties
    for (let c = 1; c <= desiredCols; c++) {
      const key = `header-${c}`;
      const cellState = (this as any).getCellState(key);
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
        const cellState = (this as any).getCellState(key);
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

    (this as any).updateCellVisuals();
    if (this.elements.createTableBtn) this.elements.createTableBtn.disabled = false;
    (this as any).markChangesForReset(); // Enable reset button after AI content is applied
    showMessage(`AI content applied (clamped to ${desiredRows} rows x ${desiredCols} cols). Review/edit then click Create Table.`, 'success');
  }

  private handleTableSelected(msg: PluginMessage): void {
    // Handle when a Data table component is selected
    this.state.hasComponent = true;
    showMessage(`Selected: Data table`, "success");
    this.elements.gridContainer.style.display = "none";
    this.elements.actionButtons.style.display = "none";
    // Trigger scan-table to analyze the selected table with a small delay
    showLoader('Scanning table...');
    setTimeout(() => {
      this.sendMessage(MESSAGE_TYPES.SCAN_TABLE);
    }, 100);
  }

  private handleComponentSelected(msg: PluginMessage): void {
    // Reset state when a new component is selected
    if (this.gridManager) {
      this.gridManager.clearGrid();
    } else {
      // Fallback if gridManager is not available
      this.state.cellProperties.clear();
      this.state.selectedCells.clear();
      this.state.currentEditingCell = null;
      this.state.sizeConfirmed = false;
    }
    this.state.tableFrameId = undefined;
    (this as any).clearChangesForReset(); // Disable reset button for new component
    console.log('[DEBUG] Reset state for new component selection');

    this.state.hasComponent = msg.isValidComponent;
    showMessage(msg.isValidComponent ? `Selected: ${msg.componentName}` : "Please select a component.", msg.isValidComponent ? "success" : "error");
    this.elements.gridContainer.style.display = "none";
    this.elements.actionButtons.style.display = "none";
    this.state.selectedComponent = {
      id: msg.componentId,
      name: msg.componentName,
      width: msg.componentWidth,
      properties: msg.properties || {},
      availableProperties: msg.availableProperties || [],
      propertyTypes: msg.propertyTypes || {}
    };
    if (msg.isValidComponent) {
      showLoader('Scanning table...');
      this.sendMessage(MESSAGE_TYPES.SCAN_TABLE);
    }
  }

  private handleSelectionCleared(msg: PluginMessage): void {
    this.state.hasComponent = false;
    // Only show error if we're not currently editing an existing table
    if (!this.state.tableFrameId) {
      showMessage("Please select a table cell component.", "error");
      this.elements.gridContainer.style.display = "none";
      this.elements.actionButtons.style.display = "none";
      this.state.selectedComponent = null;
      this.elements.landingPage.style.display = 'block';
      hideMessage(); // Hide status messages on landing page
    }
  }

  private handleTableCreated(msg: PluginMessage): void {
    hideLoader();
    // Remove loading state from button
    this.elements.createTableBtn.classList.remove('loading');
    this.elements.createTableBtn.disabled = false;

    if (msg.isComponent) {
      showMessage("Table component created successfully! You can now reuse it.", "success");
    } else {
      showMessage("Table created successfully!", "success");
    }
    // Show landing page and componentModeBtn again for new table generation
    this.elements.landingPage.style.display = 'flex';
    hideMessage(); // Hide status messages on landing page
    if (this.elements.componentModeBtn) {
      this.elements.componentModeBtn.style.display = 'inline-block';
    }
    this.elements.gridContainer.style.display = 'none';
    this.elements.actionButtons.style.display = 'none';
    this.elements.propertyEditor.style.display = 'none';
    // Clear all cell properties after table is generated
    if (this.gridManager) {
      this.gridManager.clearGrid();
    } else {
      this.state.cellProperties.clear();
    }
    // Reset button text to "Create Table" for new tables
    this.elements.createTableBtn.textContent = 'Create Table';
    this.state.tableFrameId = undefined;
  }

  private handleFakeDataResponse(msg: PluginMessage): void {
    hideLoader();
    const sample = msg.data;
    console.log('[DEBUG] Received fake-data-response', sample);
    if (!Array.isArray(sample) || sample.length === 0) {
      showMessage('No data generated.', 'error');
      return;
    }
    // Apply faker data to the correct cells
    if (this.pendingFakerContext) {
      const { mode, key, fakerMethod } = this.pendingFakerContext;
      const availableProps: string[] = this.state.selectedComponent?.availableProperties || [];
      const propertyTypes: { [key: string]: any } = this.state.selectedComponent?.propertyTypes || {};

      // Helper to find the correct TEXT property key for a cell
      const getCellTextProp = () => {
        // Prefer the first TEXT property; fallback to generic 'Cell text'
        return availableProps.find(p => propertyTypes[p] === 'TEXT') || 'Cell text';
      };

      // Helper to find the property that controls text visibility
      const getTextVisibilityProp = () => {
        // Look for a boolean prop that likely controls text visibility
        return availableProps.find(p =>
          propertyTypes[p] === 'BOOLEAN' &&
          (p.toLowerCase().includes('show') || p.toLowerCase().includes('text')) &&
          !p.toLowerCase().includes('slot') // Exclude slot toggles
        );
      };

      const applyAiData = (cellKey: string, data: string) => {
        const cellState = (this as any).getCellState(cellKey);
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
        this.state.selectedCells.add(cellKey);
      };

      if (mode === 'cell') {
        applyAiData(key, sample[0]);
      } else if (mode === 'row') {
        const [row] = key.split(',').map(Number);
        for (let c = 1; c <= this.state.gridCols; c++) {
          const k = `${row},${c}`;
          applyAiData(k, sample[(c - 1) % sample.length]);
        }
      } else if (mode === 'column') {
        const [, col] = key.split(',').map(Number);
        for (let r = 1; r <= this.state.gridRows; r++) {
          const k = `${r},${col}`;
          applyAiData(k, sample[(r - 1) % sample.length]);
        }
      }

      this.pendingFakerContext = null;
      (this as any).updateCellVisuals();
      (this as any).closePropertyEditor();
      document.querySelectorAll('.cell').forEach(cell => {
        const divCell = cell as HTMLDivElement;
        const key = `${divCell.dataset.row},${divCell.dataset.col}`;
        if (this.state.selectedCells.has(key)) {
          divCell.classList.add('selected');
        }
      });
    }
  }

  private handleWatsonxDataResponse(msg: PluginMessage): void {
    hideLoader();
    const wx = msg.data;
    if (!Array.isArray(wx) || wx.length === 0) {
      showMessage('No data generated from watsonx.ai.', 'error');
      return;
    }
    if (this.pendingFakerContext) {
      const { mode, key } = this.pendingFakerContext;
      const availableProps: string[] = this.state.selectedComponent?.availableProperties || [];
      const propertyTypes: { [key: string]: any } = this.state.selectedComponent?.propertyTypes || {};
      const getCellTextProp = () => { return availableProps.find(p => propertyTypes[p] === 'TEXT') || 'Cell text'; };
      const getTextVisibilityProp = () => { return availableProps.find(p => propertyTypes[p] === 'BOOLEAN' && (p.toLowerCase().includes('show') || p.toLowerCase().includes('text')) && !p.toLowerCase().includes('slot')); };
      const applyAiData = (cellKey: string, data: string) => {
        const cellState = (this as any).getCellState(cellKey);
        const textProp = getCellTextProp();
        const visibilityProp = getTextVisibilityProp();
        if (textProp) cellState.properties[textProp] = data;
        if (visibilityProp) cellState.properties[visibilityProp] = true;
        this.state.selectedCells.add(cellKey);
      };
      if (mode === 'cell') {
        applyAiData(key, wx[0]);
      } else if (mode === 'row') {
        const [row] = key.split(',').map(Number);
        for (let c = 1; c <= this.state.gridCols; c++) {
          const k = `${row},${c}`;
          applyAiData(k, wx[(c - 1) % wx.length]);
        }
      } else if (mode === 'column') {
        const [, col] = key.split(',').map(Number);
        for (let r = 1; r <= this.state.gridRows; r++) {
          const k = `${r},${col}`;
          applyAiData(k, wx[(r - 1) % wx.length]);
        }
      }
      this.pendingFakerContext = null;
      (this as any).updateCellVisuals();
      (this as any).closePropertyEditor();
      document.querySelectorAll('.cell').forEach(cell => {
        const divCell = cell as HTMLDivElement;
        const k = `${divCell.dataset.row},${divCell.dataset.col}`;
        if (this.state.selectedCells.has(k)) divCell.classList.add('selected');
      });
    }
  }

  private handlePromptFallbackNotice(msg: PluginMessage): void {
    hideLoader();
    showMessage(`Could not find a specific category for "${msg.prompt}". Using general text.`, "error");
  }

  private handleCreationError(msg: PluginMessage): void {
    hideLoader();
    // Remove loading state from button
    this.elements.createTableBtn.classList.remove('loading');
    this.elements.createTableBtn.disabled = false;
    showMessage(`Error: ${msg.message}`, "error");
  }

  private handleComponentProperties(msg: PluginMessage): void {
    this.state.hasComponent = true;
    this.state.componentProps = msg.props;
    this.state.selectedComponent = msg.component;
    this.state.componentWidth = msg.component.width;

    // Hide landing page and show status message
    this.elements.landingPage.style.display = 'none';
    this.elements.statusMessage.style.display = 'block';
    this.elements.statusMessage.textContent = `Selected component: ${msg.component.name}`;
    this.elements.statusMessage.className = 'status success';
    this.elements.gridContainer.style.display = 'flex';
  }

  private handleScanTable(msg: PluginMessage): void {
    console.log('[PLUGIN] Received scan-table message from UI');
    // This is typically handled by the plugin, not the UI
    // Placeholder for any UI-side scan table logic
  }

  private handleScanTableResult(msg: PluginMessage): void {
    // Implementation is quite long - will be added in a separate method
    this.handleScanTableResultImpl(msg);
  }

  private handleGeneratedTableSelectedNoSettings(msg: PluginMessage): void {
    console.log(`[UI] Generated table selected but no settings available`);
    this.state.tableFrameId = msg.tableId;
    this.state.hasComponent = true;

    // Show UI without error message
    this.elements.landingPage.style.display = 'none';
    this.elements.gridContainer.style.display = 'flex';
    this.elements.actionButtons.style.display = 'flex';

    // Reset to default state
    this.state.gridRows = 5;
    this.state.gridCols = 5;
    if (this.gridManager) {
      this.gridManager.clearGrid();
    } else {
      this.state.cellProperties.clear();
    }

    // Change button text to indicate this is an update
    this.elements.createTableBtn.textContent = "Update Table";
    this.elements.createTableBtn.disabled = false;

    showMessage("Generated table selected. You can modify and update it.", "success");
  }

  private handleEditExistingTable(msg: PluginMessage): void {
    // Implementation is quite long - will be added in a separate method
    this.handleEditExistingTableImpl(msg);
  }

  private handleTableUpdated(msg: PluginMessage): void {
    hideLoader();
    // Remove loading state from button
    this.elements.createTableBtn.classList.remove('loading');
    this.elements.createTableBtn.disabled = false;

    showMessage("Table updated successfully!", "success");
    // Keep the button as "Update Table" since we're still editing the same table
    this.elements.createTableBtn.textContent = 'Update Table';
  }

  private handleComponentInfo(msg: PluginMessage): void {
    console.log(`[MessageHandler] Received component info:`, msg.component);
    if (msg.component) {
      this.state.selectedComponent = msg.component;
      console.log(`[MessageHandler] Component info loaded, updating cell visuals with component data`);
      // Now update the visuals with the component information
      (this as any).updateCellVisuals();
    } else {
      console.log(`[MessageHandler] No component info received, using fallback for cell properties`);
      // If no component info, we'll still try to display properties from cellProperties
      (this as any).updateCellVisuals();
    }
  }

  private handleWatsonxSettings(msg: PluginMessage): void {
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

  // Helper methods for complex message handlers
  private handleScanTableResultImpl(msg: PluginMessage): void {
    console.log(`[MessageHandler] handleScanTableResult called - this will clear cell properties!`);
    console.log(`[MessageHandler] Current tableFrameId: ${this.state.tableFrameId}, hasComponent: ${this.state.hasComponent}`);
    
    hideLoader();
    if (this.elements.scanTableBtn) {
      this.elements.scanTableBtn.disabled = false;
      this.elements.scanTableBtn.textContent = 'Scan Table';
    }
    showMessage(msg.message, msg.success ? 'success' : 'error');

    if (!msg.success) return;

    // Reset state for new table generation
    if (this.gridManager) {
      this.gridManager.clearGrid();
    } else {
      // Fallback if gridManager is not available
      this.state.cellProperties.clear();
      this.state.selectedCells.clear();
      this.state.currentEditingCell = null;
      this.state.sizeConfirmed = false;
    }
    this.state.tableFrameId = undefined;
    console.log('[DEBUG] Reset state for new table generation');

    // Hide the initial landing page message
    this.elements.landingPage.style.display = 'none';

    if (this.elements.scanTableSection) {
      this.elements.scanTableSection.style.display = 'none';
    }
    const details = msg.details;
    if (details.bodyCellComponent) {
      this.state.selectedComponent = details.bodyCellComponent;
      this.state.hasComponent = true;
      // Store header/footer components for property editing
      this.state.headerCellComponent = details.headerCellComponent || null;
      this.state.footerComponent = details.footerComponent || null;
      console.log('Component template set from scan:', this.state.selectedComponent);

      // Update the grid to reflect any changes in component properties
      (this as any).renderHeaderFooterGrids?.();
    } else {
      showMessage('Could not find a body cell template in the scanned table.', 'error');
      if (this.elements.scanTableSection) this.elements.scanTableSection.style.display = 'flex';
      return;
    }
    // Setup grid
    this.state.gridCols = details.numCols || 5;
    this.state.gridRows = 5; // default
    this.gridManager?.createGrid();
    // Automatically select all cells
    document.querySelectorAll('.cell').forEach(cell => {
      const key = `${(cell as HTMLDivElement).dataset.row},${(cell as HTMLDivElement).dataset.col}`;
      this.state.selectedCells.add(key);
      cell.classList.add('selected');
    });
    this.state.sizeConfirmed = true;
    // Show grid UI with animation
    this.elements.gridContainer.style.display = 'flex';
    this.elements.actionButtons.style.display = 'flex';
    requestAnimationFrame(() => {
      this.elements.gridContainer.classList.add('show');
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
            <label style="margin-right: 8px;">Rows: <input type="number" id="scanRowsInput" class="styled-input" value="${this.state.gridRows}" min="1" style="width: 70px;"></label>
            <label>Cols: <input type="number" id="scanColsInput" class="styled-input" value="${this.state.gridCols}" min="1" style="width: 70px;"></label>
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
              <input type="checkbox" id="scanSelectableToggle" checked>
              <label for="scanSelectableToggle">Selectable</label>
            </div>
            <div class="property-field checkbox">
              <input type="checkbox" id="scanExpandableToggle" checked>
              <label for="scanExpandableToggle">Expandable</label>
            </div>
        </div>
    `;
    // Insert scan options before the grid in the grid container
    this.elements.gridContainer.insertBefore(optionsDiv, this.elements.gridContainer.firstChild);
    this.elements.scanOptionsContainer = optionsDiv;

    const updateGridFromInputs = () => {
      const newRows = parseInt((document.getElementById('scanRowsInput') as HTMLInputElement).value, 10);
      const newCols = parseInt((document.getElementById('scanColsInput') as HTMLInputElement).value, 10);
      if (newRows !== this.state.gridRows || newCols !== this.state.gridCols) {
        this.state.gridRows = newRows;
        this.state.gridCols = newCols;
        this.gridManager?.createGrid();
        // Re-select all cells after recreating grid
        document.querySelectorAll('.cell').forEach(cell => {
          const key = `${(cell as HTMLDivElement).dataset.row},${(cell as HTMLDivElement).dataset.col}`;
          this.state.selectedCells.add(key);
          cell.classList.add('selected');
        });
        // --- Filter cellProperties to keep only valid cells/headers/footers ---
        for (const key of Array.from(this.state.cellProperties.keys())) {
          // Match cell keys like "row,col"
          const match = key.match(/^([0-9]+),([0-9]+)$/);
          if (match) {
            const row = parseInt(match[1], 10);
            const col = parseInt(match[2], 10);
            if (row > this.state.gridRows || col > this.state.gridCols) {
              this.state.cellProperties.delete(key);
            }
          }
          // Remove header properties for columns that no longer exist
          if (key.startsWith('header-')) {
            const col = parseInt(key.split('-')[1], 10);
            if (col > this.state.gridCols) {
              this.state.cellProperties.delete(key);
            }
          }
          // Optionally, handle footer if you want to remove it when footer is not present
          // (No-op for now)
        }
        (this as any).updateCellVisuals?.();
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
    this.elements.createTableBtn.textContent = 'Create Table';
    this.elements.createTableBtn.disabled = false;
    // Optionally, you could hide the button if you have two separate buttons
    // this.elements.createTableBtn.style.display = 'inline-block';

    (this as any).setMode?.('edit');
    (this as any).updateCreateButtonState?.();
  }

  private handleEditExistingTableImpl(msg: PluginMessage): void {
    console.log(`[MessageHandler] Processing edit-existing-table with settings:`, msg.settings);
    console.log(`[MessageHandler] cellProperties keys:`, Object.keys(msg.settings?.cellProperties || {}));
    
    const settings = msg.settings;
    if (!settings) return;
    
    this.state.tableFrameId = msg.tableId;

    // Reset state
    this.state.selectedCells.clear();
    this.state.sizeConfirmed = false;
    this.state.currentEditingCell = null;
    this.state.hasComponent = true; // Ensure this is true for generated tables

    // Show UI with animation
    console.log(`[MessageHandler] Before showing UI - landingPage exists: ${!!this.elements.landingPage}, gridContainer exists: ${!!this.elements.gridContainer}, actionButtons exists: ${!!this.elements.actionButtons}`);
    
    if (this.elements.landingPage) this.elements.landingPage.style.display = 'none';
    if (this.elements.gridContainer) this.elements.gridContainer.style.display = 'flex';
    if (this.elements.actionButtons) this.elements.actionButtons.style.display = 'flex';
    
    requestAnimationFrame(() => {
      if (this.elements.gridContainer) this.elements.gridContainer.classList.add('show');
    });

    console.log(`[MessageHandler] After showing UI - Grid display: ${this.elements.gridContainer?.style.display}, Action buttons display: ${this.elements.actionButtons?.style.display}`);

    // Update state from settings
    this.state.gridRows = settings.rows || 5;
    this.state.gridCols = settings.columns || 5;

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
    this.state.cellProperties = convertedCellProperties;
    console.log(`[MessageHandler] Converted cell properties:`, Array.from(convertedCellProperties.keys()));

    // Create grid and update visuals
    console.log(`[MessageHandler] Creating grid - gridManager exists: ${!!this.gridManager}, rows: ${this.state.gridRows}, cols: ${this.state.gridCols}`);
    if (this.gridManager) {
      this.gridManager.createGrid();
      console.log(`[MessageHandler] Grid created successfully`);
    } else {
      console.error(`[MessageHandler] gridManager is null - cannot create grid`);
    }

    // Request component information from backend for proper display
    console.log(`[MessageHandler] Requesting component info from backend for tableId: ${this.state.tableFrameId}`);
    this.requestComponentInfo(this.state.tableFrameId);
    console.log(`[MessageHandler] Component info request sent`);

    // Update visuals after a short delay to allow component info to load
    setTimeout(() => {
      console.log(`[MessageHandler] Calling updateCellVisuals - selectedComponent exists: ${!!this.state.selectedComponent}`);
      if (this.updateCellVisuals) {
        this.updateCellVisuals();
      }
    }, 100);

    // Set sizeConfirmed to true for existing tables (they already have a confirmed size)
    this.state.sizeConfirmed = true;

    // Update button text to indicate this is an update
    if (this.elements.createTableBtn) {
      this.elements.createTableBtn.textContent = "Update Table";
      this.elements.createTableBtn.disabled = false;
    }

    showMessage("Generated table selected. You can modify and update it.", "success");
  }
}