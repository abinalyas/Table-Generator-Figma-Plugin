import type { State, Elements } from '../types';
import { showMessage, hideMessage } from '../utils/helpers';
import { HEADER_CELL_MODEL, FOOTER_MODEL } from '../utils/constants';
import { RemainingUtils } from './RemainingUtils';

/**
 * StateManager component for centralized state management
 * Handles all state transitions, UI updates, and property editor operations
 */
export class StateManager {
  private state: State;
  private elements: Elements;
  
  // External dependencies
  private gridManager: any = null;
  private propertyRenderer: any = null;
  private messageHandler: any = null;
  
  // Track if user has made any changes that can be reset
  private hasChangesToReset = false;

  constructor(state: State, elements: Elements) {
    this.state = state;
    this.elements = elements;
  }

  /**
   * Set external dependencies
   */
  setDependencies(dependencies: {
    gridManager?: any;
    propertyRenderer?: any;
    messageHandler?: any;
    renderHeaderFooterGrids?: () => void;
    createGrid?: () => void;
  }): void {
    Object.assign(this, dependencies);
  }

  /**
   * Set the application mode (selection or edit)
   */
  setMode(newMode: State['mode']): void {
    this.state.mode = newMode;
    document.body.classList.remove('selection-mode', 'edit-mode');
    document.body.classList.add(`${newMode}-mode`);

    this.state.isDragging = false;
    this.elements.gridHighlight.style.display = 'none';

    this.updateModeDependentVisibility();

    // Show colWidthInput if default mode is column
    if (this.state.applyMode === 'column' && this.elements.colWidthInput) {
      this.elements.colWidthInput.style.display = 'inline-block';
    }
  }

  /**
   * Update visibility of UI elements based on current mode and state
   */
  updateModeDependentVisibility(): void {
    // Only show action buttons if a valid component is selected and table size is confirmed
    if (this.state.hasComponent && this.state.sizeConfirmed) {
      this.elements.actionButtons.style.display = 'flex';
    } else {
      this.elements.actionButtons.style.display = 'none';
    }
  }

  /**
   * Get or create cell state for a given key
   */
  getCellState(key: string): any {
    if (!this.state.cellProperties.has(key)) {
      this.state.cellProperties.set(key, { properties: {} });
    }
    return this.state.cellProperties.get(key);
  }

  /**
   * Set cell properties for a given key
   */
  setCellProperties(key: string, properties: any): void {
    console.log(`[StateManager] setCellProperties called for ${key}:`, properties);
    console.log(`[StateManager] selectedComponent exists: ${!!this.state.selectedComponent}`);
    const cellState = this.getCellState(key);
    cellState.properties = { ...cellState.properties, ...properties };
    this.state.cellProperties.set(key, cellState);
    console.log(`[StateManager] Cell properties set, total cells with properties: ${this.state.cellProperties.size}`);
  }

  /**
   * Set column width for a cell and optionally apply to entire column
   */
  setCellColumnWidth(key: string, width: number, applyToColumn: boolean = false): void {
    const cellState = this.getCellState(key);
    cellState.colWidth = width;
    
    if (applyToColumn) {
      const [row, col] = key.includes('-') ? ['header', key.split('-')[1]] : key.split(',');
      if (key.startsWith('header-')) {
        // Apply to entire column for header cells
        const colNum = key.split('-')[1];
        const columnCells = Array.from({ length: this.state.gridRows }, (_, i) => `${i + 1},${colNum}`);
        for (const cellKey of columnCells) {
          const colCellState = this.getCellState(cellKey);
          colCellState.colWidth = width;
        }
      } else {
        // Apply to column for body cells
        const columnCells = Array.from({ length: this.state.gridRows }, (_, i) => `${i + 1},${col}`);
        for (const cellKey of columnCells) {
          const colCellState = this.getCellState(cellKey);
          colCellState.colWidth = width;
        }
      }
    }
  }

  /**
   * Set the currently editing cell
   */
  setCurrentEditingCell(key: string | null): void {
    this.state.currentEditingCell = key;
  }

  /**
   * Get the currently editing cell
   */
  getCurrentEditingCell(): string | null {
    return this.state.currentEditingCell;
  }

  /**
   * Set grid dimensions
   */
  setGridDimensions(rows: number, cols: number): void {
    this.state.gridRows = rows;
    this.state.gridCols = cols;
  }

  /**
   * Get grid dimensions
   */
  getGridDimensions(): { rows: number; cols: number } {
    return { rows: this.state.gridRows, cols: this.state.gridCols };
  }

  /**
   * Get selected component information
   */
  getSelectedComponent(): any {
    console.log('[DEBUG] StateManager getSelectedComponent called, returning:', this.state.selectedComponent);
    return this.state.selectedComponent;
  }

  /**
   * Get header cell component information
   */
  getHeaderCellComponent(): any {
    return this.state.headerCellComponent;
  }

  /**
   * Get apply mode
   */
  getApplyMode(): string {
    return this.state.applyMode;
  }

  /**
   * Set apply mode
   */
  setApplyMode(mode: 'cell' | 'row' | 'column'): void {
    this.state.applyMode = mode;
  }

  /**
   * Update the create button state based on current conditions
   */
  updateCreateButtonState(): void {
    console.log('[StateManager] updateCreateButtonState called - sizeConfirmed:', this.state.sizeConfirmed, 'button text before:', this.elements.createTableBtn?.textContent);
    this.elements.createTableBtn.disabled = !this.state.sizeConfirmed;
    this.elements.createTableBtn.style.opacity = this.state.sizeConfirmed ? '1' : '0.5';
    this.elements.createTableBtn.style.cursor = this.state.sizeConfirmed ? 'pointer' : 'not-allowed';
    this.elements.createTableBtn.title = this.state.sizeConfirmed ? 'Create table in Figma' : 'Select a component first';
    console.log('[StateManager] updateCreateButtonState done - button text after:', this.elements.createTableBtn?.textContent, 'disabled:', this.elements.createTableBtn?.disabled);
  }

  /**
   * Update the reset button state
   */
  updateResetButtonState(): void {
    const resetBtn = this.elements.clearSelectionBtn;
    if (resetBtn) {
      resetBtn.disabled = !this.hasChangesToReset;
      resetBtn.style.opacity = this.hasChangesToReset ? '1' : '0.5';
      resetBtn.style.cursor = this.hasChangesToReset ? 'pointer' : 'not-allowed';
      resetBtn.title = this.hasChangesToReset ? 'Reset all table properties' : 'No changes to reset';
    }
  }

  /**
   * Mark that changes have been made that can be reset
   */
  markChangesForReset(): void {
    this.hasChangesToReset = true;
    this.updateResetButtonState();
  }

  /**
   * Clear the changes for reset flag
   */
  clearChangesForReset(): void {
    this.hasChangesToReset = false;
    this.updateResetButtonState();
  }

  /**
   * Reset all table properties to default state
   */
  resetTableProperties(): void {
    if (this.gridManager) {
      this.gridManager.clearGrid();
    } else {
      // Fallback if gridManager is not available
      this.state.cellProperties.clear();
      this.state.selectedCells.clear();
      this.state.currentEditingCell = null;
      this.state.sizeConfirmed = false;
    }

    // Reset component selection state
    this.state.selectedComponent = null;
    this.state.headerCellComponent = null;
    this.state.footerComponent = null;
    this.state.hasComponent = false;
    this.state.tableFrameId = undefined;

    // Hide UI elements
    this.elements.gridContainer.style.display = 'none';
    this.elements.actionButtons.style.display = 'none';
    this.elements.propertyEditor.style.display = 'none';

    // Show landing page
    this.elements.landingPage.style.display = 'block';

    // Recreate the grid to clear all styling including slot colors
    this.gridManager?.createGrid();
    this.clearChangesForReset(); // Reset button should be disabled after reset
    showMessage("All cell properties have been reset.", "success");
  }

  /**
   * Open property editor for a specific cell
   */
  openPropertyEditor(key: string): void {
    // Allow property editor to open regardless of mode
    // if (this.state.mode !== 'edit') return;

    // If we don't have component info and this is a body cell, request it
    if (!this.state.selectedComponent && !key.startsWith('header-') && key !== 'footer') {
      console.log(`[openPropertyEditor] No component info available, requesting...`);
      this.messageHandler?.requestComponentInfo(this.state.tableFrameId);

      // Wait for component info to be received before proceeding (with timeout)
      let attempts = 0;
      const maxAttempts = 20; // 1 second timeout (20 * 50ms)

      const waitForComponentInfo = () => {
        attempts++;
        if (this.state.selectedComponent) {
          console.log(`[openPropertyEditor] Component info received, proceeding with editor`);
          this.openPropertyEditorInternal(key); // Call the internal function now that we have component info
        } else if (attempts < maxAttempts) {
          setTimeout(waitForComponentInfo, 50);
        } else {
          console.log(`[openPropertyEditor] Timeout waiting for component info, proceeding with fallback`);
          // Proceed with the editor anyway, using fallback logic
          this.openPropertyEditorInternal(key);
        }
      };
      waitForComponentInfo();
      return;
    }

    // Call the internal function to avoid infinite recursion
    this.openPropertyEditorInternal(key);
  }

  /**
   * Close the property editor
   */
  closePropertyEditor(): void {
    this.elements.propertyEditor.style.opacity = '0';
    this.elements.propertyEditorOverlay.style.display = 'none';
    setTimeout(() => {
      this.elements.propertyEditor.style.display = 'none';
      this.state.currentEditingCell = null;
    }, 300);
  }

  /**
   * Internal property editor opening logic
   */
  private openPropertyEditorInternal(key: string): void {
    try {
      // Guarantee the editor has required nodes
      RemainingUtils.ensurePropertyEditorScaffold(this.elements);
      // Show column width option based on current apply mode
      if (this.state.applyMode === 'cell' || this.state.applyMode === 'column') {
        this.elements.colWidthContainer.style.display = 'block';
        console.log(`[DEBUG] Initial: Showing column width option for ${this.state.applyMode} mode`);
      } else {
        this.elements.colWidthContainer.style.display = 'none';
        console.log(`[DEBUG] Initial: Hiding column width option for ${this.state.applyMode} mode`);
      }
      // Set default apply mode and update tab visual state
      if (key.startsWith('header-') || key === 'footer') {
        // For header/footer cells, default to 'cell' mode
        this.state.applyMode = 'cell';
        // Update tab visual state
        document.querySelectorAll('.apply-option').forEach(opt => opt.classList.remove('active'));
        const cellOption = document.querySelector('.apply-option[data-apply="cell"]') as HTMLElement;
        if (cellOption) {
          cellOption.classList.add('active');
        }
      } else {
        // For body cells, default to 'column' mode
        this.state.applyMode = 'column';
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

      // Apply option event listeners are now handled by EventManager
      // applyOptions.forEach((option, index) => {
      //   console.log(`[DEBUG] Attaching listener to option ${index}:`, option.textContent);
      //   // Remove existing listeners to avoid duplicates
      //   option.removeEventListener('click', handleApplyOptionClick);
      //   // Add new listener
      //   option.addEventListener('click', handleApplyOptionClick);
      //   console.log(`[DEBUG] Listener attached to:`, option.textContent);
      // });

      this.state.currentEditingCell = key;

      let availableProps: string[] = this.state.selectedComponent?.availableProperties || [];
      let propertyTypes: { [key: string]: any } = this.state.selectedComponent?.propertyTypes || {};

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
      if (secondTextPropKey && this.elements.secondLineContainer) {
        this.elements.secondLineContainer.style.display = 'none';
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
        const cellState = this.getCellState(key);
        props = cellState.properties || {};
        // Centralized static field logic
        this.propertyRenderer?.updateStaticFieldsVisibilityAndValues(availableProps, propertyTypes, props);
      }

      if (key.startsWith('header-')) {
        const cellState = this.state.cellProperties.get(key);
        props = (cellState && cellState.properties) ? cellState.properties : {};

        // Apply default values for header cell properties
        for (const fieldDef of HEADER_CELL_MODEL) {
          if (fieldDef.defaultValue && props[fieldDef.name] === undefined) {
            props[fieldDef.name] = fieldDef.defaultValue;
          }
        }

        label = `Header ${key.split('-')[1]}`;
        console.log('[DEBUG] openPropertyEditor header', { props });
        this.propertyRenderer?.renderDynamicPropertyFieldsFromModel(HEADER_CELL_MODEL, props);
      } else if (key === 'footer') {
        const cellState = this.state.cellProperties.get(key);
        props = (cellState && cellState.properties) ? cellState.properties : {};
        label = 'Footer';
        console.log('[DEBUG] openPropertyEditor footer', { props });

        // For footer, only show the "Type" property
        const footerModel = FOOTER_MODEL.filter(field => field.label === "Type");
        this.propertyRenderer?.renderDynamicPropertyFieldsFromModel(footerModel, props);
      } else {
        const cellState = this.getCellState(key);
        props = cellState.properties || {};
        const [row, col] = key.split(',');
        label = `(${row},${col})`;
        console.log('[DEBUG] openPropertyEditor body', { availableProps, propertyTypes, props });
        this.propertyRenderer?.renderDynamicPropertyFields(availableProps, propertyTypes, props);
      }

      // Show col width for all cells (body, header, footer)
      let width: number | undefined = undefined;

      if (key.startsWith('header-')) {
        // For header cells, get width from the corresponding column
        const col = key.split('-')[1];
        const colCells = Array.from({ length: this.state.gridRows }, (_, r) => `${r + 1},${col}`);
        for (const k of colCells) {
          const cellState = this.state.cellProperties.get(k);
          if (cellState && cellState.colWidth) {
            width = cellState.colWidth;
            break;
          }
        }
        if (width === undefined) {
          width = 120;
        }
        console.log('[DEBUG] Header col width UI exists?', !!this.elements.colWidthContainer, !!this.elements.colWidthInput, 'value to set:', width);
        if (this.elements.colWidthContainer) this.elements.colWidthContainer.style.display = 'block';
        if (this.elements.colWidthInput) this.elements.colWidthInput.value = String(width);
        console.log(`[DEBUG] Header cell ${key} - colWidth: ${width}, container display: ${this.elements.colWidthContainer.style.display}`);
      } else if (key === 'footer') {
        // For footer, use a default width
        width = 120;
        console.log('[DEBUG] Footer col width UI exists?', !!this.elements.colWidthContainer, !!this.elements.colWidthInput, 'value to set:', width);
        // Hide column width for footer cells
        if (this.elements.colWidthContainer) this.elements.colWidthContainer.style.display = 'none';
        if (this.elements.colWidthInput) this.elements.colWidthInput.value = String(width);
        console.log(`[DEBUG] Footer cell - colWidth: ${width}, container display: ${this.elements.colWidthContainer.style.display}`);
      } else {
        // For body cells, get width from the column
        const col = key.split(',')[1];
        const colCells = Array.from({ length: this.state.gridRows }, (_, r) => `${r + 1},${col}`);
        for (const k of colCells) {
          const cellState = this.state.cellProperties.get(k);
          if (cellState && cellState.colWidth) {
            width = cellState.colWidth;
            break;
          }
        }
        if (width === undefined) {
          width = 120;
        }
        console.log('[DEBUG] Body col width UI exists?', !!this.elements.colWidthContainer, !!this.elements.colWidthInput, 'value to set:', width, 'for key', key);
        if (this.elements.colWidthContainer) this.elements.colWidthContainer.style.display = 'none';
        if (this.elements.colWidthInput) this.elements.colWidthInput.value = String(width);
        console.log(`[DEBUG] Body cell ${key} - colWidth: ${width}, container display: ${this.elements.colWidthContainer.style.display}`);
      }

      // Set title and show editor
      this.elements.propertyEditorTitle.innerHTML = `Cell Properties <span>${label}</span>`;
      this.elements.editingCellCoords.textContent = label;
      this.elements.propertyEditorOverlay.style.display = 'block';
      this.elements.propertyEditor.style.display = 'block';

      // Final check: Ensure custom cell text input is shown for body cells if toggle is checked
      if (!key.startsWith('header-') && key !== 'footer') {
        if (this.elements.customCellTextToggle && this.elements.customCellTextToggle.checked && this.elements.customCellTextContainer) {
          this.elements.customCellTextContainer.style.display = 'block';
          console.log('[DEBUG] Final check: Setting customCellTextContainer display to block for body cell');
        }
      }

      setTimeout(() => {
        this.elements.propertyEditor.style.opacity = '1';
      }, 10);
    } catch (error) {
      console.error('Error in openPropertyEditor:', error);
      if (typeof figma !== 'undefined') {
        figma.notify('Failed to open property editor');
      }
    }
  }

  /**
   * Update cell visuals based on current properties
   */
  updateCellVisuals(): void {
    const availableProps: string[] = this.state.selectedComponent?.availableProperties || [];
    console.log(`[updateCellVisuals] availableProps:`, availableProps);
    console.log(`[updateCellVisuals] selectedComponent:`, this.state.selectedComponent);
    console.log(`[updateCellVisuals] selectedComponent.availableProperties:`, this.state.selectedComponent?.availableProperties);
    console.log(`[updateCellVisuals] Total cellProperties entries: ${this.state.cellProperties.size}`);

    // Update body cells
    document.querySelectorAll('.cell').forEach(cell => {
      const divCell = cell as HTMLDivElement;
      const key = `${divCell.dataset.row},${divCell.dataset.col}`;
      const props = this.state.cellProperties.get(key);
      console.log(`[updateCellVisuals] Cell ${key}:`, props);

      if (props && props.properties && Object.keys(props.properties).length > 0) {
        let displayText = '';
        let hasSlotEnabled = false;

        // If we have selectedComponent, use its property types
        // Skip selectedComponent if it's the wrong component (Resizer) or has wrong properties
        const isWrongComponent = this.state.selectedComponent?.name === 'Resizer' || 
                                availableProps.includes('Size') || 
                                availableProps.length === 0;
        
        if (this.state.selectedComponent && availableProps.length > 0 && !isWrongComponent) {
          // Only use selectedComponent if it has the right properties
          // Check for slot property first
          const slotProp = availableProps.find(p =>
            this.state.selectedComponent!.propertyTypes[p] === 'BOOLEAN' && p.toLowerCase().includes('slot'));
          if (slotProp && (props.properties[slotProp] === true || props.properties[slotProp] === 'true')) {
            hasSlotEnabled = true;
          }

          // Find any TEXT property to display as cell text
          for (const propName of availableProps) {
            const type = this.state.selectedComponent.propertyTypes[propName];
            if (type === 'TEXT' && typeof props.properties[propName] === 'string' && props.properties[propName].toString().trim() !== '') {
              displayText = props.properties[propName].toString();
              break; // Use the first TEXT property found
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
      for (let c = 1; c <= this.state.gridCols; c++) {
        const cell = headerGrid.querySelector(`.header-cell:nth-child(${c})`) as HTMLDivElement;
        if (cell) {
          const key = `header-${c}`;
          const props = this.state.cellProperties.get(key);
          let displayText = '';

          // Try to get text from properties
          if (props && props.properties) {
            // Use header cell component properties if available
            let textKey = '';
            if (this.state.headerCellComponent?.availableProperties && this.state.headerCellComponent?.propertyTypes) {
              const availableProps: string[] = this.state.headerCellComponent.availableProperties;
              const propertyTypes: { [key: string]: any } = this.state.headerCellComponent.propertyTypes;
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
        const props = this.state.cellProperties.get('footer');
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

  // ensurePropertyEditorScaffold moved to RemainingUtils component
}