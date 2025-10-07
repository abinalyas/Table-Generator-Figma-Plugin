// Grid management functionality
import type { State, Elements } from '../types';
import { HEADER_CELL_MODEL, FOOTER_MODEL } from '../utils/constants';
import { sortColumnData } from '../utils/sorting';
import { RemainingUtils } from './RemainingUtils';

export class GridManager {
  private state: State;
  private elements: Elements;
  private stateManager: any = null;

  constructor(state: State, elements: Elements) {
    this.state = state;
    this.elements = elements;
    
    // Ensure cellProperties is properly initialized
    if (!(this.state.cellProperties instanceof Map)) {
      console.warn('[GridManager] cellProperties is not a Map, reinitializing...');
      this.state.cellProperties = new Map();
    }
    
    // Ensure selectedCells is properly initialized
    if (!(this.state.selectedCells instanceof Set)) {
      console.warn('[GridManager] selectedCells is not a Set, reinitializing...');
      this.state.selectedCells = new Set();
    }
  }

  /**
   * Set external dependencies
   */
  setDependencies(dependencies: { 
    stateManager?: any;
    showCellTooltip?: (element: HTMLDivElement) => void;
    hideCellTooltip?: () => void;
    openPropertyEditor?: (key: string) => void;
    renderHeaderFooterGrids?: () => void;
    updateCreateButtonState?: () => void;
  }): void {
    Object.assign(this, dependencies);
  }

  /**
   * Create and render the main grid
   */
  createGrid(): void {
    RemainingUtils.createGrid(this.state, this.elements, {
      showCellTooltip: (this as any).showCellTooltip || (() => {}),
      hideCellTooltip: (this as any).hideCellTooltip || (() => {}),
      openPropertyEditor: (this as any).openPropertyEditor || (() => {}),
      renderHeaderFooterGrids: (this as any).renderHeaderFooterGrids || (() => {}),
      updateCreateButtonState: (this as any).updateCreateButtonState || (() => {})
    });
  }

  /**
   * Render header and footer grids
   */
  renderHeaderFooterGrids(): void {
    this.renderHeaderGrid();
    this.renderFooterGrid();
  }

  private renderHeaderGrid(): void {
    const headerGrid = document.getElementById('headerGrid');
    if (!headerGrid) return;

    headerGrid.innerHTML = '';
    headerGrid.style.setProperty('--header-cols', String(this.state.gridCols));

    const sortedCellProperties = sortColumnData(this.state.cellProperties, this.state.gridCols, this.state.gridRows);

    for (let col = 1; col <= this.state.gridCols; col++) {
      const cell = document.createElement('div');
      cell.className = 'header-cell';
      cell.dataset.col = String(col);

      const headerKey = `header-${col}`;
      const headerState = sortedCellProperties.get(headerKey);
      let displayText = '';

      if (headerState?.properties) {
        const textProp = this.state.headerCellComponent?.availableProperties?.find(p =>
          this.state.headerCellComponent!.propertyTypes[p] === 'TEXT'
        );
        displayText = (textProp && headerState.properties[textProp]) ||
                     headerState.properties['Cell text#12234:32'] || '';
      }

      cell.textContent = displayText || `H${col}`;
      if (displayText) {
        cell.classList.add('edited');
        cell.style.fontWeight = 'bold';
      }

      cell.addEventListener('click', () => this.handleCellClick(`header-${col}`));
      cell.addEventListener('mouseenter', (e) => this.handleCellMouseEnter(e.target as HTMLDivElement));
      cell.addEventListener('mouseleave', () => this.handleCellMouseLeave());
      
      headerGrid.appendChild(cell);
    }
  }

  private renderFooterGrid(): void {
    const footerGrid = document.getElementById('footerGrid');
    if (!footerGrid) return;

    footerGrid.innerHTML = '';
    const cell = document.createElement('div');
    cell.className = 'footer-cell';
    cell.style.gridColumn = `span ${this.state.gridCols}`;

    const footerState = this.state.cellProperties.get('footer');
    let displayText = '';

    if (footerState?.properties) {
      const textProp = this.state.footerComponent?.availableProperties?.find(p =>
        this.state.footerComponent!.propertyTypes[p] === 'TEXT'
      );
      displayText = (textProp && footerState.properties[textProp]) ||
                   footerState.properties['Total items#12006:49'] || '';
    }

    cell.textContent = displayText || 'Footer';
    if (displayText) {
      cell.classList.add('edited');
      cell.style.fontWeight = 'bold';
    }

    cell.addEventListener('click', () => this.handleCellClick('footer'));
    cell.addEventListener('mouseenter', (e) => this.handleCellMouseEnter(e.target as HTMLDivElement));
    cell.addEventListener('mouseleave', () => this.handleCellMouseLeave());
    
    footerGrid.appendChild(cell);
  }

  private handleCellClick(cellKey: string): void {
    // This will be connected to the property editor
    // For now, we'll emit a custom event that the main UI can listen to
    const event = new CustomEvent('cellClick', { detail: { cellKey } });
    document.dispatchEvent(event);
  }

  private handleCellMouseEnter(cell: HTMLDivElement): void {
    // This will be connected to tooltip functionality
    const event = new CustomEvent('cellMouseEnter', { detail: { cell } });
    document.dispatchEvent(event);
  }

  private handleCellMouseLeave(): void {
    // This will be connected to tooltip functionality
    const event = new CustomEvent('cellMouseLeave');
    document.dispatchEvent(event);
  }

  private updateModeDependentVisibility(): void {
    // Only show action buttons if a valid component is selected and table size is confirmed
    if (this.state.hasComponent && this.state.sizeConfirmed) {
      this.elements.actionButtons.style.display = 'flex';
    } else {
      this.elements.actionButtons.style.display = 'none';
    }
  }

  /**
   * Update grid size and recreate
   */
  updateGridSize(cols: number, rows: number): void {
    this.stateManager?.setGridDimensions(rows, cols);
    this.createGrid();
  }

  /**
   * Clear all cell properties and recreate grid
   */
  clearGrid(): void {
    try {
      if (this.state.cellProperties && typeof this.state.cellProperties.clear === 'function') {
        this.state.cellProperties.clear();
      }
      if (this.state.selectedCells && typeof this.state.selectedCells.clear === 'function') {
        this.state.selectedCells.clear();
      }
      this.stateManager?.setCurrentEditingCell(null);
      this.state.sizeConfirmed = false;
      this.createGrid();
    } catch (error) {
      console.error('[GridManager] Error in clearGrid:', error);
      // Reinitialize if there's an error
      this.state.cellProperties = new Map();
      this.state.selectedCells = new Set();
      this.stateManager?.setCurrentEditingCell(null);
      this.state.sizeConfirmed = false;
      this.createGrid();
    }
  }
}