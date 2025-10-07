// Property editor functionality
import type { State, Elements, PropertyDefinition } from '../types';
import { HEADER_CELL_MODEL, FOOTER_MODEL } from '../utils/constants';
import { PropertyRenderer } from './PropertyRenderer';
import { RemainingUtils } from './RemainingUtils';

export class PropertyEditor {
  private state: State;
  private elements: Elements;
  private propertyRenderer: PropertyRenderer;
  private stateManager: any = null;

  constructor(state: State, elements: Elements) {
    this.state = state;
    this.elements = elements;
    this.propertyRenderer = new PropertyRenderer(state, elements);
  }

  /**
   * Set external dependencies
   */
  setDependencies(dependencies: { stateManager?: any }): void {
    Object.assign(this, dependencies);
  }

  /**
   * Open property editor for a specific cell
   */
  openEditor(cellKey: string): void {
    this.stateManager?.setCurrentEditingCell(cellKey);
    
    // Set title and coordinates
    if (cellKey.startsWith('header-')) {
      const col = cellKey.split('-')[1];
      this.elements.propertyEditorTitle.textContent = `Header Cell ${col}`;
      this.elements.editingCellCoords.textContent = `Header Column ${col}`;
    } else if (cellKey === 'footer') {
      this.elements.propertyEditorTitle.textContent = 'Footer Cell';
      this.elements.editingCellCoords.textContent = 'Footer';
    } else {
      const [row, col] = cellKey.split(',');
      this.elements.propertyEditorTitle.textContent = `Cell ${row},${col}`;
      this.elements.editingCellCoords.textContent = `Row ${row}, Column ${col}`;
    }

    // Render appropriate property fields
    this.renderPropertyFields(cellKey);
    
    // Show the editor
    this.showEditor();
  }

  /**
   * Close the property editor
   */
  closeEditor(): void {
    this.elements.propertyEditor.style.opacity = '0';
    this.elements.propertyEditorOverlay.style.display = 'none';
    setTimeout(() => {
      this.elements.propertyEditor.style.display = 'none';
      this.stateManager?.setCurrentEditingCell(null);
    }, 300);
  }

  /**
   * Save current property values
   */
  saveProperties(): void {
    const cellKey = this.stateManager?.getCurrentEditingCell();
    if (!cellKey) return;
    
    if (cellKey.startsWith('header-')) {
      this.saveHeaderProperties(cellKey);
    } else if (cellKey === 'footer') {
      this.saveFooterProperties(cellKey);
    } else {
      this.saveCellProperties(cellKey);
    }

    this.closeEditor();
    
    // Emit event to refresh grid
    const event = new CustomEvent('propertiesSaved', { detail: { cellKey } });
    document.dispatchEvent(event);
  }

  private renderPropertyFields(cellKey: string): void {
    if (cellKey.startsWith('header-')) {
      this.renderHeaderFields(cellKey);
    } else if (cellKey === 'footer') {
      this.renderFooterFields(cellKey);
    } else {
      this.renderCellFields(cellKey);
    }
  }

  private renderHeaderFields(cellKey: string): void {
    const cellState = this.stateManager?.getCellState(cellKey);
    const props = cellState.properties || {};
    
    // Apply default values for header cell properties
    for (const fieldDef of HEADER_CELL_MODEL) {
      if (fieldDef.defaultValue && props[fieldDef.name] === undefined) {
        props[fieldDef.name] = fieldDef.defaultValue;
      }
    }
    
    this.renderDynamicPropertyFields(HEADER_CELL_MODEL, props);
  }

  private renderFooterFields(cellKey: string): void {
    const cellState = this.stateManager?.getCellState(cellKey);
    const props = cellState.properties || {};
    
    // For footer, only show the "Type" property
    const footerModel = FOOTER_MODEL.filter(field => field.label === "Type");
    this.renderDynamicPropertyFields(footerModel, props);
  }

  private renderCellFields(cellKey: string): void {
    const cellState = this.stateManager?.getCellState(cellKey);
    const props = cellState?.properties || {};
    
    console.log('[DEBUG] PropertyEditor renderCellFields for', cellKey);
    console.log('[DEBUG] PropertyEditor cellState:', cellState);
    console.log('[DEBUG] PropertyEditor props:', props);
    
    // Get available properties from the selected component
    const selectedComponent = this.stateManager?.getSelectedComponent();
    console.log('[DEBUG] PropertyEditor selectedComponent:', selectedComponent);
    const availableProps = selectedComponent?.availableProperties || [];
    const propertyTypes = selectedComponent?.propertyTypes || {};
    
    console.log('[DEBUG] renderCellFields', { availableProps, propertyTypes, props });
    
    // Convert to PropertyDefinition format for consistency
    const cellModel: PropertyDefinition[] = [];
    
    // Check if we have wrong component info (Resizer component or only Size property)
    const isWrongComponent = selectedComponent?.name === 'Resizer' || 
                            (availableProps.length === 1 && availableProps[0] === 'Size') ||
                            availableProps.length === 0;
    
    let propsToProcess = availableProps;
    let typesToUse = propertyTypes;
    
    // If wrong component info, use actual properties from cell data
    if (isWrongComponent && Object.keys(props).length > 0) {
      console.log('[DEBUG] Wrong component detected, using actual cell properties:', Object.keys(props));
      propsToProcess = Object.keys(props);
      // Create fallback property types based on property names
      typesToUse = {};
      for (const propName of propsToProcess) {
        if (propName.toLowerCase().includes('text')) {
          typesToUse[propName] = 'TEXT';
        } else if (propName.toLowerCase().includes('slot') || propName.toLowerCase().includes('show')) {
          typesToUse[propName] = 'BOOLEAN';
        } else {
          typesToUse[propName] = 'VARIANT';
        }
      }
      console.log('[DEBUG] Using fallback property types:', typesToUse);
    }
    
    for (const propName of propsToProcess) {
      // Skip Size property since it can't be updated
      if (propName === 'Size') continue;
      
      const type = typesToUse[propName];
      const label = this.cleanPropName(propName);
      
      let options: string[] | undefined;
      if (type === 'VARIANT') {
        // Define options based on property name
        if (propName.toLowerCase().includes('state')) {
          options = ['Enabled', 'Disabled', 'Focus'];
        } else if (propName.toLowerCase().includes('size')) {
          options = ['Extra large', 'Large', 'Small'];
        } else {
          options = [String(props[propName] || '')];
        }
      }
      
      cellModel.push({
        name: propName,
        label: label,
        type: type as 'TEXT' | 'VARIANT' | 'BOOLEAN',
        options: options
      });
    }
    
    this.renderDynamicPropertyFields(cellModel, props);
  }

  private renderDynamicPropertyFields(model: PropertyDefinition[], props: any): void {
    // Use PropertyRenderer for consistent field rendering
    this.propertyRenderer.renderDynamicPropertyFieldsFromModel(model, props);
  }

  // Field creation methods are now handled by PropertyRenderer

  private saveHeaderProperties(cellKey: string): void {
    const cellState = this.stateManager?.getCellState(cellKey);
    const existingProps = cellState?.properties || {};
    const newProps: any = { ...existingProps };

    for (const fieldDef of HEADER_CELL_MODEL) {
      const { name, type, defaultValue } = fieldDef;
      const input = document.getElementById(`dynamic-${name}`) as HTMLInputElement | HTMLSelectElement;
      
      if (input) {
        if (type === 'BOOLEAN') {
          newProps[name] = (input as HTMLInputElement).checked;
        } else {
          const value = input.value;
          if (!value && existingProps[name]) {
            // Keep existing value if input is empty
          } else {
            newProps[name] = value || defaultValue || '';
          }
        }
      }
    }

    // Handle column width if specified
    let colWidth: number | undefined;
    if (this.elements.colWidthInput && this.elements.colWidthInput.value) {
      colWidth = parseInt(this.elements.colWidthInput.value, 10);
    }

    this.stateManager?.setCellProperties(cellKey, newProps);
    
    if (colWidth) {
      this.stateManager?.setCellColumnWidth(cellKey, colWidth, true);
    }

    console.log('[DEBUG] saveCellProperties header', cellKey, newProps, 'colWidth:', colWidth);
  }

  private saveFooterProperties(cellKey: string): void {
    const newProps: any = {};
    
    for (const fieldDef of FOOTER_MODEL) {
      const { name, type } = fieldDef;
      const input = document.getElementById(`dynamic-${name}`) as HTMLInputElement | HTMLSelectElement;
      
      if (input) {
        newProps[name] = type === 'BOOLEAN' ? (input as HTMLInputElement).checked : input.value;
      }
    }

    this.stateManager?.setCellProperties(cellKey, newProps);
    
    console.log('[DEBUG] saveCellProperties footer', cellKey, newProps);
  }

  private saveCellProperties(cellKey: string): void {
    const cellState = this.stateManager?.getCellState(cellKey);
    const selectedComponent = this.stateManager?.getSelectedComponent();
    const availableProps = selectedComponent?.availableProperties || [];
    const propertyTypes = selectedComponent?.propertyTypes || {};
    const newProps: any = { ...cellState?.properties };
    
    // Check if we have wrong component info (same logic as PropertyRenderer)
    const isWrongComponent = (availableProps.length === 1 && availableProps[0] === 'Size') ||
                            availableProps.length === 0;
    
    let propsToProcess = availableProps;
    let typesToUse = propertyTypes;
    
    // If wrong component info, use actual properties from cell data (same as PropertyRenderer)
    if (isWrongComponent && Object.keys(cellState?.properties || {}).length > 0) {
      console.log('[DEBUG] Wrong component detected in saveCellProperties, using actual cell properties:', Object.keys(cellState.properties));
      propsToProcess = Object.keys(cellState.properties);
      // Create fallback property types based on property names
      typesToUse = {};
      for (const propName of propsToProcess) {
        if (propName.toLowerCase().includes('text')) {
          typesToUse[propName] = 'TEXT';
        } else if (propName.toLowerCase().includes('slot') || propName.toLowerCase().includes('show')) {
          typesToUse[propName] = 'BOOLEAN';
        } else {
          typesToUse[propName] = 'VARIANT';
        }
      }
      console.log('[DEBUG] Using fallback property types in saveCellProperties:', typesToUse);
    }
    
    // Use the old approach: read from static customCellText input for text properties
    const customCellTextInput = document.getElementById('customCellText') as HTMLInputElement;
    const customCellTextToggle = document.getElementById('customCellTextToggle') as HTMLInputElement;
    
    if (customCellTextToggle?.checked && customCellTextInput?.value) {
      // Find the correct property key for cell text (same logic as old code)
      let cellTextProp = propsToProcess.find((p: string) => {
        const type = typesToUse[p];
        return type === 'TEXT' && p.toLowerCase().includes('text') && !p.toLowerCase().includes('second');
      }) || propsToProcess.find((p: string) => p.toLowerCase().includes('text')) || 'Cell text#12234:16';
      
      newProps[cellTextProp] = customCellTextInput.value;
      console.log('[DEBUG] Saving custom cell text from static input:', customCellTextInput.value, 'to property:', cellTextProp);
    }
    
    // Save properties from dynamic fields (for non-text properties)
    for (const propName of propsToProcess) {
      if (propName === 'Size') continue; // Skip Size property
      
      // Skip text properties if we're using static input
      if (propName.toLowerCase().includes('text') && !propName.toLowerCase().includes('second')) {
        continue; // Use static input instead
      }
      
      const input = document.getElementById(`dynamic-${propName}`) as HTMLInputElement | HTMLSelectElement;
      console.log(`[DEBUG] Looking for input field: dynamic-${propName}, found:`, !!input, 'value:', input?.value);
      if (input) {
        const type = typesToUse[propName];
        if (type === 'BOOLEAN') {
          newProps[propName] = (input as HTMLInputElement).checked;
        } else {
          newProps[propName] = input.value;
        }
        console.log(`[DEBUG] Updated ${propName} to:`, newProps[propName]);
      }
    }
    
    // Handle column width if specified
    let colWidth: number | undefined;
    if (this.elements.colWidthInput && this.elements.colWidthInput.value) {
      colWidth = parseInt(this.elements.colWidthInput.value, 10);
    }
    
    this.stateManager?.setCellProperties(cellKey, newProps);
    
    if (colWidth) {
      const applyMode = this.stateManager?.getApplyMode();
      const applyToColumn = applyMode === 'column';
      this.stateManager?.setCellColumnWidth(cellKey, colWidth, applyToColumn);
      
      // For row mode, apply to entire row
      if (applyMode === 'row') {
        const [row] = cellKey.split(',');
        const { cols } = this.stateManager?.getGridDimensions() || { cols: 0 };
        const rowCells = Array.from({ length: cols }, (_, i) => `${row},${i + 1}`);
        for (const key of rowCells) {
          this.stateManager?.setCellColumnWidth(key, colWidth, false);
        }
      }
    }
    
    console.log('[DEBUG] saveCellProperties body', cellKey, newProps, 'colWidth:', colWidth);
  }

  private getCellState(cellKey: string): any {
    return this.stateManager?.getCellState(cellKey);
  }

  private showEditor(): void {
    this.elements.propertyEditorOverlay.style.display = 'flex';
    this.elements.propertyEditor.style.display = 'block';
    
    requestAnimationFrame(() => {
      this.elements.propertyEditor.style.opacity = '1';
    });
  }

  private cleanPropName(propName: string): string {
    return RemainingUtils.cleanPropName(propName);
  }
}