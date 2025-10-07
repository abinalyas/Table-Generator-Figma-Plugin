// Remaining utility functions for UI operations
import type { State, Elements } from '../types';
import { sortColumnData } from '../utils/sorting';

/**
 * RemainingUtils component for utility functions
 * Contains property name cleaning, grid scaffolding, and rendering utilities
 */
export class RemainingUtils {
  /**
   * Clean property name by removing hash suffix
   */
  static cleanPropName(name: string): string {
    return name.split('#')[0].trim();
  }

  /**
   * Ensure property editor scaffold exists with required DOM elements
   */
  static ensurePropertyEditorScaffold(elements: Elements): void {
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

  /**
   * Render header and footer grids with current state
   */
  static renderHeaderFooterGrids(
    state: State, 
    elements: Elements, 
    dependencies: {
      showCellTooltip: (element: HTMLDivElement) => void;
      hideCellTooltip: () => void;
      openPropertyEditor: (key: string) => void;
    }
  ): void {
    const { showCellTooltip, hideCellTooltip, openPropertyEditor } = dependencies;
    
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

  /**
   * Create the main grid with current state and properties
   */
  static createGrid(
    state: State, 
    elements: Elements, 
    dependencies: {
      showCellTooltip: (element: HTMLDivElement) => void;
      hideCellTooltip: () => void;
      openPropertyEditor: (key: string) => void;
      renderHeaderFooterGrids: () => void;
      updateCreateButtonState: () => void;
    }
  ): void {
    const { showCellTooltip, hideCellTooltip, openPropertyEditor, renderHeaderFooterGrids, updateCreateButtonState } = dependencies;
    
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
}