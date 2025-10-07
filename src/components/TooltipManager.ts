// Tooltip management functionality
import type { State } from '../types';

export class TooltipManager {
  private state: State;
  private tooltip: HTMLDivElement;
  private hoverTimeout: number | null = null;

  constructor(state: State) {
    this.state = state;
    this.tooltip = state.tooltip;
  }

  /**
   * Show tooltip for a cell
   */
  showCellTooltip(cell: HTMLDivElement): void {
    // Clear any existing timeout
    if (this.hoverTimeout) {
      clearTimeout(this.hoverTimeout);
    }

    this.hoverTimeout = window.setTimeout(() => {
      const row = cell.dataset.row;
      const col = cell.dataset.col;
      let cellKey = '';
      let cellState: any = null;

      if (cell.classList.contains('header-cell')) {
        cellKey = `header-${col}`;
        cellState = this.state.cellProperties.get(cellKey);
      } else if (cell.classList.contains('footer-cell')) {
        cellKey = 'footer';
        cellState = this.state.cellProperties.get(cellKey);
      } else if (row && col) {
        cellKey = `${row},${col}`;
        cellState = this.state.cellProperties.get(cellKey);
      }

      if (cellState && cellState.properties && Object.keys(cellState.properties).length > 0) {
        let tooltipContent = '<div class="tooltip-content">';
        
        for (const [key, value] of Object.entries(cellState.properties)) {
          const cleanKey = key.split('#')[0].trim();
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
        this.tooltip.innerHTML = tooltipContent;

        // Position tooltip
        const rect = cell.getBoundingClientRect();
        this.tooltip.style.left = `${rect.right + 10}px`;
        this.tooltip.style.top = `${rect.top}px`;
        this.tooltip.style.maxWidth = '300px';
        this.tooltip.style.display = 'block';

        // Ensure tooltip is in the DOM
        if (!this.tooltip.parentElement) {
          document.body.appendChild(this.tooltip);
        }
      } else {
        this.hideCellTooltip();
      }
    }, 300); // 300ms delay
  }

  /**
   * Hide tooltip
   */
  hideCellTooltip(): void {
    if (this.hoverTimeout) {
      clearTimeout(this.hoverTimeout);
      this.hoverTimeout = null;
    }
    this.tooltip.style.display = 'none';
  }

  /**
   * Handle cell mouse enter event
   */
  handleCellMouseEnter(cell: HTMLDivElement): void {
    this.showCellTooltip(cell);
  }

  /**
   * Handle cell mouse leave event
   */
  handleCellMouseLeave(): void {
    this.hideCellTooltip();
  }
}