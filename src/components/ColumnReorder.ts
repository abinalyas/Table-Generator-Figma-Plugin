import type { State, Elements } from '../types';

/**
 * ColumnReorder component for handling column reordering and deletion
 */
export class ColumnReorder {
  private state: State;
  private elements: Elements;
  private stateManager: any = null;
  private gridManager: any = null;

  constructor(state: State, elements: Elements) {
    this.state = state;
    this.elements = elements;
    this.setupEventListeners();
  }

  /**
   * Set external dependencies
   */
  setDependencies(dependencies: {
    stateManager?: any;
    gridManager?: any;
  }): void {
    Object.assign(this, dependencies);
  }

  private setupEventListeners(): void {
    // Make openColumnReorderModal available globally for HTML to call
    (window as any).openColumnReorderModal = () => this.openColumnReorderModal();
    console.log('[ColumnReorder] Global function set up:', typeof (window as any).openColumnReorderModal);

    // Also set up direct button listener as backup
    const reorderBtn = document.getElementById('reorderColumnsTextBtn');
    if (reorderBtn) {
      reorderBtn.addEventListener('click', () => {
        console.log('[ColumnReorder] Direct button click handler triggered');
        this.openColumnReorderModal();
      });
    }

    // Setup modal event listeners
    const closeBtn = document.getElementById('closeReorderModal');
    const cancelBtn = document.getElementById('cancelReorderBtn');
    const applyBtn = document.getElementById('applyReorderBtn');
    const overlay = document.getElementById('columnReorderOverlay');

    closeBtn?.addEventListener('click', () => this.closeModal());
    cancelBtn?.addEventListener('click', () => this.closeModal());
    applyBtn?.addEventListener('click', () => this.applyReorder());
    
    // Close modal when clicking overlay
    overlay?.addEventListener('click', (e) => {
      if (e.target === overlay) this.closeModal();
    });
  }

  private openColumnReorderModal(): void {
    console.log('[ColumnReorder] Opening reorder modal');
    const modal = document.getElementById('columnReorderOverlay');
    const columnList = document.getElementById('columnList');
    
    console.log('[ColumnReorder] Modal element:', modal);
    console.log('[ColumnReorder] Column list element:', columnList);
    console.log('[ColumnReorder] Grid cols:', this.state.gridCols);
    
    if (!modal || !columnList) {
      console.error('[ColumnReorder] Modal or column list not found!');
      return;
    }

    // Clear existing content
    columnList.innerHTML = '';

    // Create column items
    for (let col = 1; col <= this.state.gridCols; col++) {
      const columnItem = document.createElement('div');
      columnItem.className = 'column-item';
      columnItem.draggable = true;
      columnItem.dataset.column = col.toString();
      
      // Get header text for this column
      const headerKey = `header-${col}`;
      const headerProps = this.state.cellProperties.get(headerKey);
      const headerText = this.getColumnHeaderText(headerProps) || `Column ${col}`;
      
      columnItem.innerHTML = `
        <div class="column-content">
          <span class="drag-handle">⋮⋮</span>
          <span class="column-name">${headerText}</span>
          <button class="column-delete-btn" data-column="${col}" title="Delete column">×</button>
        </div>
      `;

      // Add drag event listeners
      this.addDragListeners(columnItem);
      
      // Add delete listener
      const deleteBtn = columnItem.querySelector('.column-delete-btn');
      deleteBtn?.addEventListener('click', (e) => {
        e.stopPropagation();
        this.deleteColumn(col);
      });

      columnList.appendChild(columnItem);
    }

    console.log('[ColumnReorder] Created', this.state.gridCols, 'column items');
    console.log('[ColumnReorder] Setting modal display to flex');
    
    // Show modal with proper animation
    modal.style.display = 'flex';
    console.log('[ColumnReorder] Modal display set, adding show class');
    
    requestAnimationFrame(() => {
      modal.classList.add('show');
      console.log('[ColumnReorder] Show class added, modal should be visible');
    });
  }

  private getColumnHeaderText(headerProps: any): string {
    if (!headerProps?.properties) return '';
    
    // Look for text properties in header
    const textProps = Object.keys(headerProps.properties).filter(prop =>
      prop.toLowerCase().includes('text') && 
      typeof headerProps.properties[prop] === 'string' &&
      headerProps.properties[prop].trim() !== ''
    );
    
    return textProps.length > 0 ? headerProps.properties[textProps[0]] : '';
  }

  private addDragListeners(item: HTMLElement): void {
    item.addEventListener('dragstart', (e) => {
      item.classList.add('dragging');
      if (e.dataTransfer) {
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/html', item.outerHTML);
        e.dataTransfer.setData('text/plain', item.dataset.column || '');
      }
    });

    item.addEventListener('dragend', () => {
      item.classList.remove('dragging');
    });

    item.addEventListener('dragover', (e) => {
      e.preventDefault();
      if (e.dataTransfer) e.dataTransfer.dropEffect = 'move';
    });

    item.addEventListener('drop', (e) => {
      e.preventDefault();
      const draggedColumn = e.dataTransfer?.getData('text/plain');
      const targetColumn = item.dataset.column;
      
      if (draggedColumn && targetColumn && draggedColumn !== targetColumn) {
        this.reorderColumns(parseInt(draggedColumn), parseInt(targetColumn));
      }
    });
  }

  private reorderColumns(fromCol: number, toCol: number): void {
    console.log(`[ColumnReorder] Reordering column ${fromCol} to position ${toCol}`);
    console.log(`[ColumnReorder] Current cell properties size: ${this.state.cellProperties.size}`);
    
    // Create new cell properties map with reordered columns
    const newCellProperties = new Map();
    
    // Create column mapping
    const columnMapping = new Map();
    let newColIndex = 1;
    
    // Build the new column order
    for (let col = 1; col <= this.state.gridCols; col++) {
      if (col === toCol) {
        columnMapping.set(fromCol, newColIndex++);
      }
      if (col !== fromCol) {
        columnMapping.set(col, newColIndex++);
      }
    }
    
    console.log(`[ColumnReorder] Column mapping:`, Array.from(columnMapping.entries()));
    
    // Reorder all cell properties
    for (const [key, value] of this.state.cellProperties.entries()) {
      if (key.startsWith('header-')) {
        const oldCol = parseInt(key.split('-')[1]);
        const newCol = columnMapping.get(oldCol);
        if (newCol) {
          newCellProperties.set(`header-${newCol}`, value);
        }
      } else if (key.includes(',')) {
        const [row, oldCol] = key.split(',').map(Number);
        const newCol = columnMapping.get(oldCol);
        if (newCol) {
          newCellProperties.set(`${row},${newCol}`, value);
        }
      } else {
        newCellProperties.set(key, value);
      }
    }
    
    console.log(`[ColumnReorder] New cell properties size: ${newCellProperties.size}`);
    this.state.cellProperties = newCellProperties;
    this.refreshModalContent();
  }

  private deleteColumn(colToDelete: number): void {
    console.log(`[ColumnReorder] Deleting column ${colToDelete}`);
    console.log(`[ColumnReorder] Current grid cols: ${this.state.gridCols}`);
    console.log(`[ColumnReorder] Current cell properties size: ${this.state.cellProperties.size}`);
    
    if (this.state.gridCols <= 1) {
      alert('Cannot delete the last column');
      return;
    }
    
    // Create new cell properties map without the deleted column
    const newCellProperties = new Map();
    
    // Create column mapping (shift columns after deleted one)
    const columnMapping = new Map();
    let newColIndex = 1;
    
    for (let col = 1; col <= this.state.gridCols; col++) {
      if (col !== colToDelete) {
        columnMapping.set(col, newColIndex++);
      }
    }
    
    console.log(`[ColumnReorder] Column mapping:`, Array.from(columnMapping.entries()));
    
    // Remove deleted column and reindex remaining columns
    for (const [key, value] of this.state.cellProperties.entries()) {
      if (key.startsWith('header-')) {
        const oldCol = parseInt(key.split('-')[1]);
        if (oldCol !== colToDelete) {
          const newCol = columnMapping.get(oldCol);
          if (newCol) {
            newCellProperties.set(`header-${newCol}`, value);
          }
        }
      } else if (key.includes(',')) {
        const [row, oldCol] = key.split(',').map(Number);
        if (oldCol !== colToDelete) {
          const newCol = columnMapping.get(oldCol);
          if (newCol) {
            newCellProperties.set(`${row},${newCol}`, value);
          }
        }
      } else {
        newCellProperties.set(key, value);
      }
    }
    
    console.log(`[ColumnReorder] Old cell properties size: ${this.state.cellProperties.size}`);
    console.log(`[ColumnReorder] New cell properties size: ${newCellProperties.size}`);
    
    this.state.cellProperties = newCellProperties;
    this.state.gridCols--;
    
    console.log(`[ColumnReorder] Updated grid cols to: ${this.state.gridCols}`);
    
    // Update grid size input
    const colsInput = document.getElementById('scanColsInput') as HTMLInputElement;
    if (colsInput) colsInput.value = this.state.gridCols.toString();
    
    // Refresh modal content to show updated columns
    this.refreshModalContent();
  }

  private refreshModalContent(): void {
    // Update modal content without closing it
    const columnList = document.getElementById('columnList');
    if (!columnList) return;

    console.log('[ColumnReorder] Refreshing modal content');
    
    // Clear existing content
    columnList.innerHTML = '';

    // Create column items with updated data
    for (let col = 1; col <= this.state.gridCols; col++) {
      const columnItem = document.createElement('div');
      columnItem.className = 'column-item';
      columnItem.draggable = true;
      columnItem.dataset.column = col.toString();
      
      // Get header text for this column
      const headerKey = `header-${col}`;
      const headerProps = this.state.cellProperties.get(headerKey);
      const headerText = this.getColumnHeaderText(headerProps) || `Column ${col}`;
      
      columnItem.innerHTML = `
        <div class="column-content">
          <span class="drag-handle">⋮⋮</span>
          <span class="column-name">${headerText}</span>
          <button class="column-delete-btn" data-column="${col}" title="Delete column">×</button>
        </div>
      `;

      // Add drag event listeners
      this.addDragListeners(columnItem);
      
      // Add delete listener
      const deleteBtn = columnItem.querySelector('.column-delete-btn');
      deleteBtn?.addEventListener('click', (e) => {
        e.stopPropagation();
        this.deleteColumn(col);
      });

      columnList.appendChild(columnItem);
    }

    console.log('[ColumnReorder] Modal content refreshed with', this.state.gridCols, 'columns');
  }

  private applyReorder(): void {
    console.log('[ColumnReorder] Applying reorder changes');
    console.log('[ColumnReorder] Current grid cols:', this.state.gridCols);
    console.log('[ColumnReorder] Cell properties size:', this.state.cellProperties.size);
    
    // Update grid size input to reflect any column deletions
    const colsInput = document.getElementById('scanColsInput') as HTMLInputElement;
    if (colsInput) {
      colsInput.value = this.state.gridCols.toString();
      console.log('[ColumnReorder] Updated cols input to:', this.state.gridCols);
    }
    
    // Recreate the grid with new column order
    if (this.gridManager) {
      console.log('[ColumnReorder] Calling gridManager.createGrid()');
      this.gridManager.createGrid();
    } else {
      console.error('[ColumnReorder] GridManager not available!');
    }
    
    // Update cell visuals
    if (this.stateManager) {
      console.log('[ColumnReorder] Calling stateManager methods');
      this.stateManager.updateCellVisuals();
      this.stateManager.markChangesForReset();
    } else {
      console.error('[ColumnReorder] StateManager not available!');
    }
    
    console.log('[ColumnReorder] Apply completed, closing modal');
    this.closeModal();
  }

  private closeModal(): void {
    const modal = document.getElementById('columnReorderOverlay');
    if (modal) {
      modal.classList.remove('show');
      setTimeout(() => {
        modal.style.display = 'none';
      }, 250); // Wait for animation to complete
    }
  }
}